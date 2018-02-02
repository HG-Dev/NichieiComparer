''' Main executable for NichieiComparer '''

import os
import logging
from NichieiComparer import ExcelDoc

def main():
    """ Main body for starting up and terminating Tweetfeeder bot """
    # pylint: disable=no-member
    os.environ['MECAB_CHARSET'] = 'UTF-8'
    logging.basicConfig(level=logging.DEBUG)
    try:
        translated_doc = ExcelDoc('./Samples/subtitle_translation.xlsx', translated=True)
        untranslated_doc = ExcelDoc('./Samples/subtitle_untranslated.xlsx', translated=False)
    except KeyboardInterrupt:
        print("Forced exit")

if __name__ == "__main__":
    main()
