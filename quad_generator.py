class QuadGenerator:
    def __init__(self):
        self.quads = []
        self.temp_counter = 0
        
    def new_temp(self):
        self.temp_counter += 1
        return f't{self.temp_counter}'
        
    def emit(self, op, arg1, arg2, result):
        """Add a new quadruple"""
        quad = (op, arg1, arg2, result)
        self.quads.append(quad)
        return result
        
    def print_quads(self):
        """Display all generated quadruples"""
        print("\nQuadruples générés:")
        print("Format: (opération, arg1, arg2, résultat)")
        print("-" * 50)
        # for i, (op, arg1, arg2, result) in enumerate(self.quads):
        #     if op in {'+', '-', '*', '/', '^'}:
        #         print(f"{i}: {result} := {arg1} {op} {arg2}")
        #     elif op == 'load':
        #         print(f"{i}: {result} := load({arg1})")
        #     elif op == 'store':
        #         print(f"{i}: store {arg1} -> {result}")
        #     elif op == 'call':
        #         print(f"{i}: {result} := call {arg1}({arg2})")
        #     else:
        #         print(f"{i}: ({op}, {arg1}, {arg2}, {result})")
        for i, (op, arg1, arg2, result) in enumerate(self.quads):
            print(f"{i}: ({op}, {arg1}, {arg2}, {result})")
        print("-" * 50)