from docx import Document
from googletrans import Translator
import time

def translate_text(text, dest_lang='hi'):
    """
    Translate the given text to the destination language.
    This function uses googletrans. For production use, consider robust error handling.
    """
    translator = Translator()
    # Googletrans may hit rate limits, so a slight delay can help
    time.sleep(0.3)
    try:
        translated = translator.translate(text, dest=dest_lang)
        return translated.text
    except Exception as e:
        print(f"Error translating text: {text}\n{e}")
        return text  # Fallback to original if translation fails

