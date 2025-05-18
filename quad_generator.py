class QuadGenerator:
    def __init__(self):
        self.quads = []
        self.temp_counter = 0
        
    def new_temp(self):
        self.temp_counter += 1
        return f't{self.temp_counter}'
        
    def emit(self, op, arg1, arg2, result):
        """Add a new quadruple"""
        if isinstance(arg1, list):
            arg1 = ', '.join(map(str, arg1))
        if isinstance(arg2, list):
            arg2 = ', '.join(map(str, arg2))
            
        quad = (op, arg1, arg2, result)
        self.quads.append(quad)
        return result
        
    def print_quads(self):
        """Display all generated quadruples"""
        print("\nQuadruples générés:")
        print("Format: (opération, arg1, arg2, résultat)")
        print("-" * 50)
        
        for i, (op, arg1, arg2, result) in enumerate(self.quads):
            print(f"{i}: ({op}, {arg1}, {arg2}, {result})")
        print("-" * 50)