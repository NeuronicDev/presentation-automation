import re
import logging
from typing import List, Dict, Optional, Union
from bs4 import BeautifulSoup


def extract_feedback_from_email(email_body: str) -> List[Dict[str, Union[str, int, None]]]:
    feedback_list: List[Dict[str, Union[str, int, None]]] = []
    ### empty for now, will implement later
    return feedback_list