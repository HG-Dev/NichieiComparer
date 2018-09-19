''' Main executable for NichieiComparer '''

import os
import logging
from NichieiComparer import ExcelDoc

def main():
    """ Main body for starting up and terminating program """
    # pylint: disable=no-member
    os.environ['MECAB_CHARSET'] = 'UTF-8'
    logging.basicConfig(level=logging.DEBUG)
    try:
        translated_doc = ExcelDoc('./Samples/subtitle_translation.xlsx', translated=True)
        untranslated_doc = ExcelDoc('./Samples/subtitle_untranslated.xlsx', translated=False)
        term_overlap = untranslated_doc.map_matching_tokens(translated_doc)
        for key, value in term_overlap.items():
            token = untranslated_doc.tokens[key].token + ":" + translated_doc.tokens[value[0]].token
            print(f"{key}:\t{value} --\t{token}")
    except KeyboardInterrupt:
        print("Forced exit")

if __name__ == "__main__":
    main()
