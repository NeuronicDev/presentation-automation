import logging
import docker
import json
import os
import tempfile
import shutil
import datetime
from typing import List, Dict, Any, Tuple
from config.config import DOCKER_IMAGE_NAME, OUTPUT_DIR

logging.basicConfig(level=logging.INFO, format='%(asctime)s - HOST - %(levelname)s - %(message)s')


DOCKERFILE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
CONTAINER_TIMEOUT = 600


def _get_docker_client() -> docker.DockerClient | None:
    try:
        client = docker.from_env()
        client.ping()
        logging.info("Successfully connected to Docker daemon.")
        return client
    except docker.errors.DockerException as e:
        logging.error(f"Failed to connect to Docker daemon: {e}")
        logging.error("Please ensure Docker is installed, running, and the current user has permissions.")
        return None
    except Exception as e:
         logging.error(f"An unexpected error occurred while connecting to Docker: {e}")
         return None


def build_docker_image() -> bool:
    client = _get_docker_client()
    if not client:
        return False

    try:
        try:
            client.images.get(DOCKER_IMAGE_NAME)
            logging.info(f"Docker image '{DOCKER_IMAGE_NAME}' already exists.")
            return True
        except docker.errors.ImageNotFound:
            logging.info(f"Docker image '{DOCKER_IMAGE_NAME}' not found. Building from {DOCKERFILE_DIR}...")
            dockerfile_path = os.path.join(DOCKERFILE_DIR, 'Dockerfile')
            if not os.path.exists(dockerfile_path):
                logging.error(f"Dockerfile not found at expected path: {dockerfile_path}")
                return False

            # Build the image
            build_logs = client.api.build(
                path=DOCKERFILE_DIR,
                dockerfile='Dockerfile',
                tag=DOCKER_IMAGE_NAME,
                rm=True, # Remove intermediate containers
                decode=True # Decode JSON stream
            )

            for chunk in build_logs:
                if 'stream' in chunk:
                    line = chunk['stream'].strip()
                    if line: 
                        logging.info(f"[BUILD] {line}")
                if 'errorDetail' in chunk:
                     logging.error(f"[BUILD ERROR] {chunk['errorDetail']['message']}")
                     return False 
                 
            logging.info(f"Docker image '{DOCKER_IMAGE_NAME}' built successfully.")
            return True
        except docker.errors.BuildError as build_err:
             logging.error(f"Docker image build failed: {build_err}")
             return False
        except docker.errors.APIError as api_err:
             logging.error(f"Docker API error checking/building image: {api_err}")
             return False

    finally:
        if client:
            client.close()



def execute_code_in_docker(tasks_with_code: List[Dict[str, Any]], original_pptx_path: str) -> Tuple[bool, str | None, List[Dict[str, Any]]]:

    if not build_docker_image():
        return False, None, {"status": "failed", "errors": [{"error": "Docker image build failed or Docker not available."}]}

    tasks_to_execute = [t for t in tasks_with_code if t.get("generated_code") and not t.get("error")]
    if not tasks_to_execute:
         logging.warning("No valid code snippets provided to execute.")
         return True, original_pptx_path, {"status": "no_tasks_to_execute", "errors": [], "processed_count": 0, "success_count": 0}

    temp_dir = None
    client = _get_docker_client()
    container = None
    final_output_path_host = None 
    execution_report = {"status": "unknown", "errors": [], "processed_count": 0, "success_count": 0}

    if not client:
         return False, None, {"status": "failed", "errors": [{"error": "Failed to connect to Docker."}]}

    try:
        temp_dir = tempfile.mkdtemp(prefix="ppt_exec_")
        working_pptx_host_path = os.path.join(temp_dir, "presentation_workcopy.pptx")

        container_pptx_path = "/app/presentation.pptx"

        shutil.copy2(original_pptx_path, working_pptx_host_path)
        logging.info(f"Copied presentation to temporary path: {working_pptx_host_path}")

        input_json = json.dumps(tasks_to_execute, indent=None) 
        input_bytes = input_json.encode('utf-8')

        logging.info(f"Running container from image '{DOCKER_IMAGE_NAME}'...")
        container = client.containers.run(
            image=DOCKER_IMAGE_NAME,
            volumes={
                working_pptx_host_path: {'bind': container_pptx_path, 'mode': 'rw'}
            },
            environment={"TASKS_INPUT": input_json},
            tty=True, 
            stdin_open=True, 
            stdout=True, 
            stderr=True,     
            detach=True,     
            # command="python executor_script.py"  # Ensure the script reads from stdin
        )
        # # Send input data directly via exec_run
        # container.reload()  # Ensure the container is ready
        # container.exec_run(f"echo '{input_json}' | python executor_script.py", stdin=True)
        
        # socket = container.attach_socket(params={'stdin': 1, 'stream': 1, 'stdout': 0, 'stderr': 0}) 
        # # socket._sock.sendall(input_bytes)
        # # socket._sock.shutdown(1)
        # print(dir(socket))
        # socket.sendall(input_bytes)
        # socket.shutdown(1)
        # socket.close()
        
        logging.info(f"Sent {len(input_bytes)} bytes of task data to container {container.short_id}.")

        logging.info(f"Waiting for container {container.short_id} to complete (timeout: {CONTAINER_TIMEOUT}s)...")
        result = container.wait(timeout=CONTAINER_TIMEOUT)
        container_status_code = result.get('StatusCode', -1)
        logging.info(f"Container finished with status code: {container_status_code}")

        stdout_logs = container.logs(stdout=True, stderr=False).decode('utf-8', errors='ignore').strip()
        stderr_logs = container.logs(stdout=False, stderr=True).decode('utf-8', errors='ignore').strip()

        if stderr_logs:
            logging.warning(f"Container {container.short_id} stderr:\n{stderr_logs}")

        overall_success = False
        if container_status_code == 0 and stdout_logs:
            try:
                log_lines = stdout_logs.split('\n')
                json_line = log_lines[-1]

                execution_report = json.loads(json_line)
                container_status = execution_report.get("status", "unknown")
                logging.info(f"Container execution status: {container_status}")
                
                if container_status in ["success", "partial_success"]:
                    overall_success = True
                    base_filename = os.path.splitext(os.path.basename(original_pptx_path))[0]
                    final_output_filename = f"{base_filename}_modified.pptx"
                    final_output_path_host = os.path.join(OUTPUT_DIR, final_output_filename) 

                    shutil.copy2(working_pptx_host_path, final_output_path_host)
                    logging.info(f"Copied modified presentation to: {final_output_path_host}")
                else:
                     logging.error(f"Container reported failure status: {container_status}")

            except json.JSONDecodeError as json_err:
                 logging.error(f"Failed to parse final status JSON from container stdout: {json_err}")
                 logging.error(f"Container stdout was:\n{stdout_logs}")
                 execution_report["errors"].append({"error": "Failed to parse container status output.", "details": str(json_err)})
        elif not stdout_logs:
             logging.error(f"Container exited with status {container_status_code} but produced no stdout report.")
             execution_report = {"status": "failed", "errors": [{"error": f"Container exited {container_status_code} with no output."}]}
        else:
            logging.error(f"Container exited with non-zero status code: {container_status_code}")
            execution_report = {"status": "failed", "errors": [{"error": f"Container exited with status {container_status_code}.", "stderr": stderr_logs}]}

        return overall_success, final_output_path_host, execution_report

    except docker.errors.NotFound as nf_err:
         logging.error(f"Docker error: {nf_err}. Container or image not found?")
         return False, None, {"status": "failed", "errors": [{"error": f"Docker resource not found: {nf_err}"}]}
    except docker.errors.APIError as api_err:
         logging.error(f"Docker API error during container execution: {api_err}")
         return False, None, {"status": "failed", "errors": [{"error": f"Docker API error: {api_err}"}]}
    except Exception as e:
        logging.error(f"Unexpected error during Docker execution process: {e}", exc_info=True)
        return False, None, {"status": "failed", "errors": [{"error": f"Host-side execution error: {e}"}]}
    
    finally:
        #Clean up container and temporary directory
        if container:
            try:
                logging.debug(f"Removing container {container.short_id}...")
                container.remove(force=True)
                logging.debug(f"Removed container {container.short_id}.")
            except docker.errors.NotFound:
                 logging.debug(f"Container {container.short_id} already removed.")
            except docker.errors.APIError as rm_err:
                logging.warning(f"Could not remove container {container.short_id}: {rm_err}")
        if client:
            client.close()
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
                logging.info(f"Cleaned up temporary directory: {temp_dir}")
            except Exception as cleanup_err:
                logging.warning(f"Failed to clean up temporary directory {temp_dir}: {cleanup_err}")