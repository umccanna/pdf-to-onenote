import os
from dotenv import load_dotenv

# Load the environment variables from .env file
load_dotenv()

__configuration = {
    "Default": {
        "InputPdfFile": "C:\\src\\data\\pdf-to-onenote\\2nd Round V2 Revised Project Narrative 052224 for Milliman.pdf",
        "OutputFolder": "C:\\src\\data\\pdf-to-onenote\\2nd Round V2 Revised Project Narrative 052224 With Tables",
        "NormalizeText": False,
        "NewSectionRegex": '^(?:I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII|XIV|XV|XVI|XVII|XVIII|XIX|XX)\\.(?!.*[\\(\\)]).+',
        #"NewSectionRegex": '^[A-Z]\\.\\s[^\\.\\d]+',
        "ExtractBasedOnTableOfContents": False,
        #"TableOfContentsPageRange": [2,4]
    }
}

def get_config(key: str):
    environment_value = os.getenv(key)
    if environment_value != None:
        return environment_value
    value = __configuration["Default"][key] if key in __configuration["Default"] else None
    return value