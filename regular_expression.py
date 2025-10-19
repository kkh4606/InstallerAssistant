import re


text_to_search = "abcjkfdlajlfjabc"

sentence = "Start a sentence and then end Start bring it to an end"


pattern = re.compile(r"^Start")

matches = pattern.finditer(sentence)

for match in matches:
    print(match)
