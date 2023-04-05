import re

text = 'The quick brown fox'

# create a capture group for the word 'quick'
pattern = r'The (quick) brown fox'

# search for the pattern in the text
match = re.search(pattern, text)

# access the captured group and assign it to a variable
captured_group = match[1]

print(captured_group)  # output: 'quick'
