import ply.lex as lex

tokens = (
    "CELL_REF", "NUMBER", "PLUS", "MINUS", "TIMES", "DIVIDE", "POWER",
    "LPAREN", "RPAREN", "COMMA", "COLON",
    "SUM", "AVERAGE", "COUNT", "MAX", "MIN", "UNIQUE",
    "TODAY", "NOW", "YEAR", "MONTH", "DAY",
    "CONCATENATE", "LEFT", "RIGHT", "MID", "LEN", "LOWER", "UPPER", "TRIM", "STRING",
    "FIND", "SEARCH", "REPLACE", "SUBSTITUTE", "TEXT", "VALUE", "PROPER", "REPT", "EXACT", "CHAR", "CODE"
)


t_PLUS   = r'\+'
t_MINUS  = r'-'
t_TIMES  = r'\*'
t_DIVIDE = r'/'
t_POWER  = r'\^'
t_LPAREN = r'\('
t_RPAREN = r'\)'
t_COMMA  = r','
t_COLON  = r':'

def t_CELL_REF(t):
    r'[A-Z]{1,2}[1-9][0-9]*'
    return t

def t_NUMBER(t):
    r'\d+(\.\d+)?'
    t.value = float(t.value) if '.' in t.value else int(t.value)
    return t

def t_STRING(t):
    r'"[^"]*"'
    t.value = t.value[1:-1]  # Remove the quotes
    return t


def t_FUNC(t):
    r'SUM|AVERAGE|COUNT|MAX|MIN|UNIQUE|TODAY|NOW|YEAR|MONTH|DAY|CONCATENATE|LEFT|RIGHT|MID|LEN|LOWER|UPPER|TRIM|FIND|SEARCH|REPLACE|SUBSTITUTE|TEXT|VALUE|PROPER|REPT|EXACT|CHAR|CODE'
    t.type = t.value.upper()  
    return t

t_ignore = " \t"

def t_error(t):
    print(f"Caractère illégal : '{t.value[0]}'")
    t.lexer.skip(1)

lexer = lex.lex()