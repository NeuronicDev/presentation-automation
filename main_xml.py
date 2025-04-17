# import os, logging, json, time, datetime, asyncio, sys, pathlib, subprocess, tempfile, shutil
# import logging.handlers
# from pptx import Presentation
# from feedback_parsing.feedback_extraction_notes import extract_feedback_from_ppt_notes
# from feedback_parsing.feedback_extraction_onslide import extract_feedback_from_ppt_onslide
# from feedback_parsing.feedback_extraction_mail import extract_feedback_from_email
# from feedback_parsing.feedback_classifier import classify_feedback_instructions
# from agents.formatting_agent import formatting_agent
# from agents.cleanup_agent import cleanup_agent
# from agents.visual_enhancement_agent import visual_enhancement_agent
# from utils.utils import generate_slide_context, convert_pptx_to_pdf, extract_slide_xml_from_ppt
# from code_manipulation.xml_code_generator import generate_modified_xml_code
# from code_manipulation.xml_code_injector import inject_xml_into_ppt

# from dotenv import load_dotenv
# load_dotenv()


# def setup_logging(log_file):
#     logger = logging.getLogger()
#     logger.setLevel(logging.INFO)
#     logger.handlers.clear()
    
#     file_handler = logging.handlers.RotatingFileHandler(log_file, encoding="utf-8")
#     file_handler.setLevel(logging.INFO)
        
#     console_handler = logging.StreamHandler()
#     console_handler.setLevel(logging.INFO)
    
#     formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
#     file_handler.setFormatter(formatter)
#     console_handler.setFormatter(formatter)
    
#     logger.addHandler(file_handler)
#     logger.addHandler(console_handler)
    
# setup_logging("wns_ppt_XML_logs2.log")


# def main(pptx_path):
    
#     email_path = "path/to/email.txt" 

#     try:
#         logging.info("Loading presentation...")
#         prs = Presentation(pptx_path)

#         pdf_path = convert_pptx_to_pdf(pptx_path, output_dir = os.path.abspath("./input_ppts/converted_pdfs"))

#         image_cache = {}
#         slide_context_cache = {}        
#         current_slide_xml = {}

#         logging.info("Extracting feedback from all sources...")
#         feedback_notes = extract_feedback_from_ppt_notes(pptx_path)
#         feedback_onslide = extract_feedback_from_ppt_onslide(pptx_path)
#         feedback_mail = extract_feedback_from_email(email_path)
        
#         all_feedback = feedback_notes + feedback_onslide + feedback_mail
#         logging.info(f"Total feedback instructions extracted: {len(all_feedback)}")

#         logging.info("Classifying and extracting tasks from feedback...")
#         categorized_tasks = classify_feedback_instructions(all_feedback)
#         logging.info(f"Categorized Tasks: {categorized_tasks}")
        
#         ########## copy of original ppt for modification
#         temp_dir = tempfile.mkdtemp(prefix="ppt_xml_")
#         working_pptx_path = os.path.join(temp_dir, "working_copy.pptx")
#         shutil.copy2(pptx_path, working_pptx_path)
        

#         logging.info("Processing Tasks with Agents & Context ---")
#         slide_tasks = {}
#         task_specifications = []
#         for task in categorized_tasks:
#             category = task["category"]
#             slide_number = task["slide_number"]
#             instruction = task["original_instruction"]
            
#             # Generate or retrieve slide context
#             if slide_number not in slide_context_cache:
#                 slide_context_cache[slide_number] = generate_slide_context(prs, slide_number, pdf_path, image_cache)
#             slide_context = slide_context_cache[slide_number]

#             # Delegate to appropriate agent
#             if category == "formatting":
#                 task_with_desc = formatting_agent(task, slide_context)
#             elif category == "cleanup":
#                 task_with_desc = cleanup_agent(task, slide_context)
#             elif category == "visual_enhancement":
#                 task_with_desc = visual_enhancement_agent(task, slide_context)
#             else:
#                 logging.warning(f"Unknown category: {category} for instruction: '{instruction}'")
#                 continue
            
#             if task_with_desc:
#                 task_specifications.extend(task_with_desc)

#             # Organize subtasks by slide number
#             if task_specifications:
#                 for subtask in task_specifications:
#                     subtask_slide = subtask.get("slide_number", slide_number)
#                     if subtask_slide not in slide_tasks:
#                         slide_tasks[subtask_slide] = []
#                     slide_tasks[subtask_slide].append(subtask)
        
                
#         # Now process tasks slide by slide, sequentially applying changes
#         for slide_number, tasks in slide_tasks.items():
#             logging.info(f"Processing {len(tasks)} tasks for slide {slide_number}")

#             current_xml = extract_slide_xml_from_ppt(working_pptx_path, slide_number)
#             if not current_xml:
#                 logging.error(f"Failed to extract XML for slide {slide_number}. Skipping all tasks for this slide.")
#                 continue
                
#             slide_context = slide_context_cache[slide_number]

#             for i, task in enumerate(tasks):
#                 task_desc = task.get("task_description", f"Task #{i+1}")
#                 logging.info(f"Applying task {i+1}/{len(tasks)} for slide {slide_number}: {task_desc}")
                
#                 modified_xml = generate_modified_xml_code(
#                     original_xml=current_xml, 
#                     agent_task_specification=task,
#                     slide_context=slide_context
#                 )
                
#                 if not modified_xml:
#                     logging.error(f"Failed to generate modified XML for task: {task_desc}")
#                     continue
                    
#                 if modified_xml == current_xml:
#                     logging.warning(f"No XML changes made for task: {task_desc}")
#                     continue
                
#                 # Update the current XML state with the modifications
#                 current_xml = modified_xml
                
#                 # Inject the modified XML back into the working PPTX
#                 success = inject_xml_into_ppt(working_pptx_path, slide_number, modified_xml)
                
#                 if success:
#                     logging.info(f"Successfully applied XML changes for task")
#                 else:
#                     logging.error(f"Failed to inject modified XML for task")
#                     break
        
  
#         base_filename = os.path.splitext(os.path.basename(pptx_path))[0]
#         final_output_filename = f"{base_filename}_xml_modified.pptx"
#         output_dir = os.path.abspath("./output_ppts")
#         os.makedirs(output_dir, exist_ok=True)
#         final_output_path = os.path.join(output_dir, final_output_filename)
        
#         shutil.copy2(working_pptx_path, final_output_path)
#         logging.info(f"Final modified presentation saved to: {final_output_path}")

#         # Cleanup temp directory
#         shutil.rmtree(temp_dir)
        

#     except FileNotFoundError as fnf_err:
#         logging.error(f"File not found error during pipeline: {fnf_err}")
#     except Exception as e:
#         logging.error(f"Pipeline failed with an unexpected error: {e}", exc_info=True)

        
# if __name__ == "__main__":
#     # pptx_path = os.path.abspath("./input_ppts/pptx/font_test1.pptx")
#     pptx_path = os.path.abspath("./input_ppts/pptx/font_test2.pptx")
#     # pptx_path = os.path.abspath("./input_ppts/pptx/cleanup_test.pptx")
#     # pptx_path = os.path.abspath("./input_ppts/pptx/cleanup_test2.pptx")
#     # pptx_path = os.path.abspath("./input_ppts/pptx/table_alignment_test.pptx")
#     # pptx_path = os.path.abspath("./input_ppts/pptx/table_alignment_test2.pptx")
#     # pptx_path = os.path.abspath("./input_ppts/pptx/consistent.pptx")
#     # pptx_path = os.path.abspath("./input_ppts/pptx/consistent3.pptx")

#     logging.info(f"input pptx: {pptx_path}")
#     main(pptx_path)
    