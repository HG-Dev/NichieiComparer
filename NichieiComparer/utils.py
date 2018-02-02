''' File and language utilities for NichieiComparer '''

import json
import re

class FileIO:
    """Collection of static methods for getting stuff out of files."""
    @staticmethod
    def get_json_dict(filepath):
        """Returns the entire JSON dict in a given file."""
        with open(filepath, encoding="utf8") as infile:
            return json.load(infile)

    @staticmethod
    def save_json_dict(filepath, dictionary):
        """Saves a JSON dict, overwriting or creating a given file."""
        with open(filepath, 'w', encoding="utf8") as outfile:
            json.dump(dictionary, outfile, ensure_ascii=False, indent=4)

class CollectionUtils:
    """Collection of static methods that hide logic operations for collections."""
    @staticmethod
    def n_highest_indices(num_list: list, n_int: int):
        """Returns a list of indices from a list with the highest to lowest
        values up to n."""
        assert n_int < len(num_list)
        return [
            x[0] for x in [
                y for y in reversed(sorted(enumerate(num_list, start=1), key=lambda i: i[1]))
            ][:n_int]
        ]

    @staticmethod
    def flatten(nested, ltypes=(list, tuple)):
        """ Reforms the nested parameter into a list of presumably hashable items. """
        nested = list(nested) # Ensure compatibility with len, etc.
        i = 0
        while i < len(nested):
            # If the object in nested at i is still a collection:
            while isinstance(nested[i], ltypes):
                # Remove empty slots
                if not nested[i]:
                    nested.pop(i)
                    i -= 1
                    break
                else:
                    # Apparently, by using a slice, we insert the entire list in-step
                    nested[i:i + 1] = nested[i]
            i += 1
        return list(nested)

    @staticmethod
    def keyword_in_string(keywords, target_str):
        """Returns true if a keyword in a list of keywords is found in the target string."""
        for keyword in keywords:
            if keyword in target_str:
                return True
        return False

    @staticmethod
    def total_adjacent_values(values):
        """Returns the number of values in a collection that are
        adjacent, i.e. plus or minus one of each other."""
        total = 0
        prev_value = None
        for value in sorted(values):
            if isinstance(prev_value, type(None)):
                prev_value = value
            else:
                if value - prev_value == 1:
                    total += 1
                    prev_value = value
        return total            

class LangUtils:
    """Collection of static methods that hide logic operations for language."""
    @staticmethod
    def is_japanese(text: str):
        """Runs a regex pattern over a string to find out if it can be said to be Japanese.
        Returns True, False, or None.
        """
        if not text:
            return None
        txt_str = str(text)
        found = re.sub('[A-Za-z0-9,.!?]+', '', txt_str)
        percent_japanese = (len(found)/len(txt_str))
        return found is not None and percent_japanese > 0.5

    @staticmethod
    def is_useful_term_jp(word_features):
        """Hides the logic of determining whether a word is worth making a key for."""
        # part_of_speech, subclass_1, subclass_2, subclass_3, inflection, conjugation, root, reading, pronunciation
        word_type = word_features[0]
        subclass = word_features[1]
        inflection = word_features[4]
        pronunciation = word_features[8]
        if CollectionUtils.keyword_in_string(['非自立'], subclass): # Remove little uninteresting bits
            return False
        if '名詞' in word_type:
            return True
            #return keyword_in_string(['一般', '代名詞', '固有名詞', '変接'], subclass)
        elif '副詞' in word_type or '助動詞' in word_type:
            return len(pronunciation) > 2
        elif '動詞' in word_type:
            return CollectionUtils.keyword_in_string(['一段', '五段'], inflection) and not '接尾' in subclass
        return False

class ExcelUtils:
    """Collection of static methods that make worksheet operations easier."""

    ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    @staticmethod
    def col_cipher(col_idx):
        """Returns an integer given a base-26 letter sequence
        or a base-26 letter sequence given an integer (index 1).
        """
        if not col_idx:
            raise TypeError()
        elif isinstance(col_idx, str):
            if len(col_idx) == 1:
                return ExcelUtils.ALPHABET.find(col_idx) + 1
            else:
                return (
                    (ExcelUtils.ALPHABET.find(col_idx[0])+1) * pow(26, len(col_idx)-1) +
                    ExcelUtils.col_cipher(col_idx[1:])
                )
        else:
            if col_idx < 1:
                return ""
            else:
                return (
                    ExcelUtils.col_cipher(col_idx / 26) +
                    ExcelUtils.ALPHABET[(int(col_idx) - 1) % 26]
                )
