import sys

def check_brackets(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    stack = []
    lines = content.split('\n')
    
    # Very naive check for {}, (), [] outside strings/comments
    # This is hard to do perfectly without a parser.
    # Let's just do a basic count of {, }, (, ), [, ]
    
    counts = {'{': 0, '}': 0, '(': 0, ')': 0, '[': 0, ']': 0}
    for char in content:
        if char in counts:
            counts[char] += 1
            
    print("Counts:", counts)
    if counts['{'] != counts['}']:
        print("Mismatched {}")
    if counts['('] != counts[')']:
        print("Mismatched ()")
    if counts['['] != counts[']']:
        print("Mismatched []")

check_brackets('e:/Mas/Aplikasi Web Bank Soal - ChatGPT/app.js')
