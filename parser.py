import ply.yacc as yacc
from lexer import tokens
import openpyxl
from datetime import datetime

# Load values from an Excel file
def load_excel_values(file_path, sheet_name="Sheet1"):
    cell_values = {}
    workbook = openpyxl.load_workbook(file_path, data_only=True)  # Read calculated values
    sheet = workbook[sheet_name]

    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
        for cell in row:
            if cell.value is not None:
                cell_values[cell.coordinate] = cell.value  # Example: {"A1": 10, "B2": 5}
    
    return cell_values

# Load the Excel file
cell_values = load_excel_values("data.xlsx")

# Operator precedence
precedence = (
    ('left', 'PLUS', 'MINUS'),
    ('left', 'TIMES', 'DIVIDE'),
    ('right', 'POWER')
)

# Rules for arithmetic operations
def p_expression_binop(p):
    """expression : expression PLUS expression
                  | expression MINUS expression
                  | expression TIMES expression
                  | expression DIVIDE expression
                  | expression POWER expression"""
    if p[2] == '+':
        p[0] = p[1] + p[3]
    elif p[2] == '-':
        p[0] = p[1] - p[3]
    elif p[2] == '*':
        p[0] = p[1] * p[3]
    elif p[2] == '/':
        p[0] = p[1] / p[3] if p[3] != 0 else print("error: Cannot divide by zero")
    elif p[2] == '^':
        p[0] = p[1] ** p[3]

def p_expression_parens(p):
    "expression : LPAREN expression RPAREN"
    p[0] = p[2]

def p_expression_number(p):
    "expression : NUMBER"
    p[0] = p[1]

def p_expression_string(p):
    "expression : STRING"
    p[0] = p[1]

# Read an Excel cell
def p_expression_cell(p):
    "expression : CELL_REF"
    cell = p[1].upper()
    p[0] = cell_values.get(cell, 0)

# Read a range of cells
def p_range(p):
    """range : CELL_REF COLON CELL_REF
             | CELL_REF"""
    if len(p) == 4:  # Range (e.g., A1:A5)
        start, end = p[1].upper(), p[3].upper()
        
        start_col, start_row = start[0], int(start[1:])
        end_col, end_row = end[0], int(end[1:])
        
        if start_col == end_col:  # Single column
            p[0] = [cell_values.get(f"{start_col}{i}", 0) for i in range(start_row, end_row + 1)]
        else:
            print("Multi-column ranges are not supported yet.")
            p[0] = []
    else:
        p[0] = [cell_values.get(p[1].upper(), 0)]

# Handle Excel functions
# def p_expression_function(p):
#     """expression : SUM LPAREN arguments RPAREN
#                   | AVERAGE LPAREN arguments RPAREN
#                   | COUNT LPAREN arguments RPAREN
#                   | MAX LPAREN arguments RPAREN
#                   | MIN LPAREN arguments RPAREN
#                   | UNIQUE LPAREN arguments RPAREN
#                   | TODAY LPAREN RPAREN
#                   | NOW LPAREN RPAREN
#                   | YEAR LPAREN expression RPAREN
#                   | MONTH LPAREN expression RPAREN
#                   | DAY LPAREN expression RPAREN
#                   | CONCATENATE LPAREN arguments RPAREN
#                   | LEFT LPAREN expression COMMA NUMBER RPAREN
#                   | RIGHT LPAREN expression COMMA NUMBER RPAREN
#                   | MID LPAREN expression COMMA NUMBER COMMA NUMBER RPAREN
#                   | LEN LPAREN expression RPAREN
#                   | LOWER LPAREN expression RPAREN
#                   | UPPER LPAREN expression RPAREN
#                   | TRIM LPAREN expression RPAREN"""
#     func = p[1].lower()
#     if func in ["sum", "average", "count", "max", "min", "unique"]:
#         values = p[3]  # List of values
#         if func == "sum":
#             p[0] = sum(values)
#         elif func == "average":
#             p[0] = sum(values) / len(values) if values else 0
#         elif func == "count":
#             p[0] = len(values)
#         elif func == "max":
#             p[0] = max(values) if values else None
#         elif func == "min":
#             p[0] = min(values) if values else None
#         elif func == "unique":
#             p[0] = list(set(values))
#     elif func == "today":
#         p[0] = datetime.today().date()
#     elif func == "now":
#         p[0] = datetime.now()
#     elif func == "year":
#         p[0] = p[3].year if isinstance(p[3], datetime) else None
#     elif func == "month":
#         p[0] = p[3].month if isinstance(p[3], datetime) else None
#     elif func == "day":
#         p[0] = p[3].day if isinstance(p[3], datetime) else None
#     elif func == "concatenate":
#         p[0] = ''.join(str(arg) for arg in p[3])
#     elif func == "left":
#         p[0] = p[3][:int(p[5])]
#     elif func == "right":
#         p[0] = p[3][-int(p[5]):]
#     elif func == "mid":
#         p[0] = p[3][int(p[5])-1:int(p[5])-1+int(p[7])]
#     elif func == "len":
#         p[0] = len(p[3])
#     elif func == "lower":
#         p[0] = p[3].lower()
#     elif func == "upper":
#         p[0] = p[3].upper()
#     elif func == "trim":
#         p[0] = p[3].strip()

def p_expression_function(p):
    """expression : SUM LPAREN arguments RPAREN
                  | AVERAGE LPAREN arguments RPAREN
                  | COUNT LPAREN arguments RPAREN
                  | MAX LPAREN arguments RPAREN
                  | MIN LPAREN arguments RPAREN
                  | UNIQUE LPAREN arguments RPAREN
                  | TODAY LPAREN RPAREN
                  | NOW LPAREN RPAREN
                  | YEAR LPAREN expression RPAREN
                  | MONTH LPAREN expression RPAREN
                  | DAY LPAREN expression RPAREN
                  | CONCATENATE LPAREN arguments RPAREN
                  | LEFT LPAREN expression COMMA NUMBER RPAREN
                  | RIGHT LPAREN expression COMMA NUMBER RPAREN
                  | MID LPAREN expression COMMA NUMBER COMMA NUMBER RPAREN
                  | LEN LPAREN expression RPAREN
                  | LOWER LPAREN expression RPAREN
                  | UPPER LPAREN expression RPAREN
                  | TRIM LPAREN expression RPAREN
                  | FIND LPAREN expression COMMA expression COMMA NUMBER RPAREN
                  | SEARCH LPAREN expression COMMA expression COMMA NUMBER RPAREN
                  | REPLACE LPAREN expression COMMA NUMBER COMMA NUMBER COMMA expression RPAREN
                  | SUBSTITUTE LPAREN expression COMMA expression COMMA expression COMMA NUMBER RPAREN
                  | TEXT LPAREN expression COMMA STRING RPAREN
                  | VALUE LPAREN expression RPAREN
                  | PROPER LPAREN expression RPAREN
                  | REPT LPAREN expression COMMA NUMBER RPAREN
                  | EXACT LPAREN expression COMMA expression RPAREN
                  | CHAR LPAREN NUMBER RPAREN
                  | CODE LPAREN expression RPAREN"""
    func = p[1].lower()
    if func in ["sum", "average", "count", "max", "min", "unique"]:
        values = p[3]  # List of values
        if func == "sum":
            p[0] = sum(values)
        elif func == "average":
            p[0] = sum(values) / len(values) if values else 0
        elif func == "count":
            p[0] = len(values)
        elif func == "max":
            p[0] = max(values) if values else None
        elif func == "min":
            p[0] = min(values) if values else None
        elif func == "unique":
            p[0] = list(set(values))
    elif func == "today":
        p[0] = datetime.today().date()
    elif func == "now":
        p[0] = datetime.now()
    elif func in ["year", "month", "day"]:
        if isinstance(p[3], str):
            try:
                date_obj = datetime.strptime(p[3], "%Y-%m-%d")
            except ValueError:
                print("Error: Invalid date format. Use YYYY-MM-DD.")
                p[0] = None
                return
        elif isinstance(p[3], datetime):
            date_obj = p[3]
        else:
            print("Error: Invalid date input.")
            p[0] = None
            return

        if func == "year":
            p[0] = date_obj.year
        elif func == "month":
            p[0] = date_obj.month
        elif func == "day":
            p[0] = date_obj.day
    elif func == "concatenate":
        p[0] = ''.join(str(arg) for arg in p[3])
    elif func == "left":
        p[0] = p[3][:int(p[5])]
    elif func == "right":
        p[0] = p[3][-int(p[5]):]
    elif func == "mid":
        p[0] = p[3][int(p[5])-1:int(p[5])-1+int(p[7])]
    elif func == "len":
        p[0] = len(p[3])
    elif func == "lower":
        p[0] = p[3].lower()
    elif func == "upper":
        p[0] = p[3].upper()
    elif func == "trim":
        p[0] = p[3].strip()
    elif func == "find":
        # FIND is case-sensitive
        find_text = p[3]
        within_text = p[5]
        start_num = int(p[7]) - 1  # Convert to 0-based index
        pos = within_text.find(find_text, start_num)
        p[0] = pos + 1 if pos != -1 else None  # Return 1-based index
    elif func == "search":
        # SEARCH is case-insensitive
        find_text = p[3].lower()
        within_text = p[5].lower()
        start_num = int(p[7]) - 1  # Convert to 0-based index
        pos = within_text.find(find_text, start_num)
        p[0] = pos + 1 if pos != -1 else None  # Return 1-based index
    elif func == "replace":
        p[0] = p[3][:int(p[5])-1] + p[9] + p[3][int(p[5])-1+int(p[7]):]
    elif func == "substitute":
        text, old_text, new_text = p[3], p[5], p[7]
        instance_num = p[9] if len(p) > 10 else None
        if instance_num:
            parts = text.split(old_text)
            if len(parts) > instance_num:
                p[0] = old_text.join(parts[:instance_num]) + new_text + old_text.join(parts[instance_num:])
            else:
                p[0] = text
        else:
            p[0] = text.replace(old_text, new_text)
    elif func == "text":
        p[0] = format(p[3], p[5])
    elif func == "value":
        try:
            p[0] = float(p[3])
        except ValueError:
            p[0] = None
    elif func == "proper":
        p[0] = p[3].title()
    elif func == "rept":
        p[0] = p[3] * int(p[5])
    elif func == "exact":
        p[0] = p[3] == p[5]
    elif func == "char":
        p[0] = chr(int(p[3]))
    elif func == "code":
        p[0] = ord(p[3][0]) if p[3] else None

# Function arguments: numbers, cells, ranges, and nested expressions
def p_arguments(p):
    """arguments : arguments COMMA expression
                 | range
                 | expression"""
    if len(p) == 4:
        p[0] = p[1] + [p[3]]
    else:
        p[0] = p[1] if isinstance(p[1], list) else [p[1]]

# Handle syntax errors
def p_error(p):
    print("Syntax error!")

# Build the parser
parser = yacc.yacc()

# Interactive testing
if __name__ == "__main__":

    while True : 
        print("""--- Menu --- \n 
              1. Display supported functions \n 
              2. Enter an Excel expression \n 
              3. Exit \n""")
        choice = int(input("Enter your choice: "))
        if choice == 1:
            print("""Supported functions:
                  \n--- Arithmetic Functions ---
                  \n- SUM(range): Calculates the sum of values in a range.
                  \n- AVERAGE(range): Computes the average of values in a range.
                  \n- COUNT(range): Counts the number of values in a range.
                  \n- MAX(range): Finds the maximum value in a range.
                  \n- MIN(range): Finds the minimum value in a range.
                  \n- UNIQUE(range): Extracts unique values from a range.
                  \n--- Date and Time Functions ---
                  \n- TODAY(): Returns today's date.
                  \n- NOW(): Returns the current date and time.
                  \n- YEAR(date): Extracts the year from a given date.
                  \n- MONTH(date): Extracts the month from a given date.
                  \n- DAY(date): Extracts the day from a given date.
                  \n--- Text Functions ---
                  \n- CONCATENATE(text1, text2, ...): Concatenates text strings.
                  \n- LEFT(text, num_chars): Extracts the left part of a text string.
                  \n- RIGHT(text, num_chars): Extracts the right part of a text string.
                  \n- MID(text, start_num, num_chars): Extracts a substring from the middle of a text string.
                  \n- LEN(text): Returns the length of a text string.
                  \n- LOWER(text): Converts text to lowercase.
                  \n- UPPER(text): Converts text to uppercase.
                  \n- TRIM(text): Removes extra spaces from text.
                  \n- FIND(find_text, within_text, [start_num]): Finds the position of a text within another text (case-sensitive).
                  \n- SEARCH(find_text, within_text, [start_num]): Finds the position of a text within another text (case-insensitive).
                  \n- REPLACE(old_text, start_num, num_chars, new_text): Replaces part of a text with another text.
                  \n- SUBSTITUTE(text, old_text, new_text, [instance_num]): Replaces specific text within a string.
                  \n- TEXT(value, format_text): Converts a value to text with a specified format.
                  \n- VALUE(text): Converts text to a numeric value.
                  \n- PROPER(text): Converts text to proper case (first letter of each word capitalized).
                  \n- REPT(text, number_times): Repeats text a specified number of times.
                  \n- EXACT(text1, text2): Compares two text strings to check if they are exactly the same (case-sensitive).
                  \n- CHAR(number): Returns the character corresponding to an ASCII code.
                  \n- CODE(text): Returns the ASCII code of the first character in a text string.""")
        elif choice == 2:
            while True:
                try:
                    expr = input("Enter an Excel expression: ").strip()
                    if expr.startswith("="):
                        expr = expr[1:]
                    result = parser.parse(expr)
                    print("Result:", result)
                except EOFError:
                    break
        elif choice == 3:
            print("Exiting...")
            break
        else:
            print("Invalid choice. Please try again.")



# Test expressions for all supported functions:
# =SUM(1, 2, 3, 4) → 10
# =AVERAGE(1, 2, 3, 4) → 2.5
# =COUNT(1, 2, 3, 4) → 4
# =MAX(1, 2, 3, 4) → 4
# =MIN(1, 2, 3, 4) → 1
# =UNIQUE(1, 2, 2, 3, 4, 4) → [1, 2, 3, 4]
# =TODAY() → Current date (e.g., 2023-10-05)
# =NOW() → Current date and time (e.g., 2023-10-05 14:30:00)
# =YEAR("2023-10-05") → 2023
# =MONTH("2023-10-05") → 10
# =DAY("2023-10-05") → 5
# =CONCATENATE("Hello", " ", "World") → "Hello World"
# =LEFT("Hello World", 5) → "Hello"
# =RIGHT("Hello World", 5) → "World"
# =MID("Hello World", 7, 5) → "World"
# =LEN("Hello World") → 11
# =LOWER("Hello World") → "hello world"
# =UPPER("Hello World") → "HELLO WORLD"
# =TRIM("  Hello World  ") → "Hello World"
# =FIND("World", "Hello World", 1) → 7                                 
# =SEARCH("world", "Hello World", 1) → 7                               
# =REPLACE("Hello World", 7, 5, "Universe") → "Hello Universe"
# =VALUE("123.45") → 123.45
# =PROPER("hello world") → "Hello World"
# =REPT("a", 5) → "aaaaa"
# =EXACT("Hello", "hello") → False
# =CHAR(65) → "A"
# =CODE("A") → 65
# =A1 → Value of cell A1
# =1/0 → Error: Cannot divide by zero
# =SUM(1, 2, AVERAGE(3, 5)) → 7
# =UPPER(LEFT("Hello World", 5)) → "HELLO"
# =YEAR(TODAY()) → Current year (e.g., 2023)
# =SUM(A1:A5) → Sum of values from A1 to A5
# =UNIQUE(A1:A5) → Unique values from A1 to A5
# =VALUE("123") + 1 → 124
# =TEXT(123, "0000") → "0123"