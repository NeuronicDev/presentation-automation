import logging
import os
from typing import List, Dict, Any, Tuple
from code_manipulation.code_corrector import generate_code_with_retry

logging.basicConfig(level=logging.INFO, format='%(asctime)s - EXECUTOR - %(levelname)s - %(message)s')

MAX_RETRIES = 3


def execute_officejs_code_tasks(tasks_with_code: List[Dict[str, Any]]) -> Tuple[bool, List[Dict[str, Any]], Dict[str, Any]]:
    logging.info("Starting Office.js task preparation for frontend execution...")

    if not tasks_with_code:
        logging.warning("No tasks provided.")
        return False, [], {
            "status": "no_tasks",
            "processed_count": 0,
            "success_count": 0,
            "errors": []
        }

    updated_tasks = []
    report = {
        "status": "unknown",
        "processed_count": 0,
        "success_count": 0,
        "errors": []
    }

    for i, task in enumerate(tasks_with_code):
        task["task_index"] = i
        code = task.get("generated_code")

        if not code:
            task["error"] = "Missing Office.js code."
            report["errors"].append({"task_index": i, "error": "Missing code."})
            updated_tasks.append(task)
            continue

        retries = 0
        success = False

        while retries < MAX_RETRIES:
            # Simulate "execution" by checking if code is valid string (basic check)
            if isinstance(code, str) and "Office.onReady" in code:
                task["error"] = None
                report["success_count"] += 1
                success = True
                break
            else:
                error_msg = "Invalid Office.js code format."
                logging.error(f"Task {i} failed: {error_msg}")
                task["error"] = error_msg

                corrected_code = generate_code_with_retry(task, error_msg, task)
                if corrected_code:
                    code = corrected_code
                    task["generated_code"] = corrected_code
                    retries += 1
                    logging.info(f"Retrying task {i} with corrected code (attempt {retries})")
                else:
                    logging.error(f"Code correction failed for task {i}")
                    break

        if not success:
            report["errors"].append({"task_index": i, "error": task.get("error", "Unknown error")})

        updated_tasks.append(task)
        report["processed_count"] += 1

    if report["success_count"] == len(updated_tasks):
        report["status"] = "success"
    elif report["success_count"] > 0:
        report["status"] = "partial_success"
    else:
        report["status"] = "failed"

    return True, updated_tasks, report
