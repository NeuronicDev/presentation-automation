import os, logging, json, time, datetime, asyncio, sys, pathlib, subprocess
import logging.handlers
from pptx import Presentation
from feedback_parsing.feedback_extraction_notes import extract_feedback_from_ppt_notes
from feedback_parsing.feedback_extraction_onslide import extract_feedback_from_ppt_onslide
from feedback_parsing.feedback_extraction_mail import extract_feedback_from_email
from feedback_parsing.feedback_classifier import classify_feedback_instructions
from agents.formatting_agent import formatting_agent
from agents.cleanup_agent import cleanup_agent
from agents.visual_enhancement_agent import visual_enhancement_agent
from utils.utils import generate_slide_context
from code_manipulation.code_generator import generate_python_code
from code_manipulation.code_executor import execute_code_in_docker
from config.config import LIBREOFFICE_PATH

from dotenv import load_dotenv
load_dotenv()


def setup_logging(log_file):
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    
    file_handler = logging.handlers.RotatingFileHandler(log_file, encoding="utf-8")
    file_handler.setLevel(logging.INFO)
        
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
setup_logging("wns_ppt_logs.log")


def main(pptx_path):
    
    email_path = "path/to/email.txt" 

    try:
        logging.info("Loading presentation...")
        prs = Presentation(pptx_path)

        pdf_path = pptx_path.replace(".pptx", ".pdf")
        logging.info(f"Converting {pptx_path} to PDF at {pdf_path}...")
        try:
            subprocess.run(
                [LIBREOFFICE_PATH, "--headless", "--convert-to", "pdf", pptx_path, "--outdir", os.path.dirname(pptx_path)],
                check=True
            )
            logging.info(f"Successfully converted {pptx_path} to {pdf_path}")
        except subprocess.CalledProcessError as e:
            logging.error(f"Failed to convert PPTX to PDF: {e}")
            sys.exit(1)


        image_cache = {}
        slide_context_cache = {}

        # Step 1: Feedback Extraction
        logging.info("Extracting feedback from all sources...")
        feedback_notes = extract_feedback_from_ppt_notes(pptx_path)
        feedback_onslide = extract_feedback_from_ppt_onslide(pptx_path)
        feedback_mail = extract_feedback_from_email(email_path)
        
        all_feedback = feedback_notes + feedback_onslide + feedback_mail
        logging.info(f"Total feedback instructions extracted: {len(all_feedback)}")

        # Step 2: Instruction Interpretation
        logging.info("Classifying and extracting tasks from feedback...")
        categorized_tasks = classify_feedback_instructions(all_feedback)
        logging.info(f"Categorized Tasks: {categorized_tasks}")

        # Step 3: Delegate tasks to specialized agents
        logging.info("Processing Tasks with Agents & Context ---")
        task_specifications = []
        for task in categorized_tasks:
            category = task["category"]
            slide_number = task["slide_number"]
            instruction = task["original_instruction"]
            
            # Generate or retrieve slide context
            if slide_number not in slide_context_cache:
                slide_context_cache[slide_number] = generate_slide_context(prs, slide_number, pdf_path, image_cache)
            slide_context = slide_context_cache[slide_number]
            
            # Delegate to appropriate agent
            if category == "formatting":
                task_with_desc = formatting_agent(task, slide_context)
            elif category == "cleanup":
                task_with_desc = cleanup_agent(task, slide_context)
            elif category == "visual_enhancement":
                task_with_desc = visual_enhancement_agent(task, slide_context)
            else:
                logging.warning(f"Unknown category: {category} for instruction: '{instruction}'")
                continue
            
            if task_with_desc:
                task_specifications.extend(task_with_desc)
    
        # Step 4: Code Generation
        logging.info("Generating Python code for task specifications...")
        tasks_with_code  = []
        for task_specification in task_specifications:
            slide_number = task_specification["slide_number"]
            slide_context = slide_context_cache[slide_number]
            code = generate_python_code(task_specification, slide_context)
            if code:
                tasks_with_code .append({
                    "slide_number": slide_number,
                    "generated_code": code,
                    "original_instruction": task_specification["original_instruction"],
                    "description": task_specification["task_description"]
                })
                logging.info(f"Generated code for task: {task_specification['task_description']}")
            else:
                logging.warning(f"Failed to generate code for task: {task_specification['task_description']}")

        logging.info(f"Generated {len(tasks_with_code )} code snippets for the given task specifications.")

        # Step 5: Safe Code Execution ---
        logging.info("--- Step 5: Executing Code in Docker ---")
        execution_success, final_output_file_path, execution_report = execute_code_in_docker(tasks_with_code, pptx_path)

        if execution_success:
            logging.info(f"Code execution completed. Status: {execution_report.get('status', 'unknown')}")
            if final_output_file_path:
                    logging.info(f"Modified presentation saved to: {final_output_file_path}")
            else: 
                    logging.error("Execution reported success but no output file path was returned.")
                    execution_success = False 

            if execution_report.get("errors"):
                    logging.warning(f"Some code snippets failed during execution ({len(execution_report['errors'])}):")
                    for error_detail in execution_report["errors"]:
                        logging.warning(f"  - Task Index {error_detail.get('task_index', 'N/A')}: {error_detail.get('error', 'Unknown error')}")
        else:
            logging.error("Code execution in Docker failed.")
            if execution_report.get("errors"):
                    logging.error("Details of execution failure:")
                    for error_detail in execution_report["errors"]:
                        logging.error(f"  - Task Index {error_detail.get('task_index', 'N/A')}: {error_detail.get('error', 'Unknown error')}")

    except FileNotFoundError as fnf_err:
            logging.error(f"File not found error during pipeline: {fnf_err}")
    except Exception as e:
        logging.error(f"Pipeline failed with an unexpected error: {e}", exc_info=True)


if __name__ == "__main__":
    pptx_path = os.path.abspath("font_test1.pptx")
    # pptx_path = os.path.abspath("cleanup_test.pptx")

    logging.info(f"input pptx: {pptx_path}")
    main(pptx_path)
    