"""
Enhanced LaTeX processor with better MathML conversion
"""

import re
import subprocess
import tempfile
import os
from pathlib import Path

try:
    from latex2mathml.converter import convert
    LATEX2MATHML_AVAILABLE = True
except ImportError:
    LATEX2MATHML_AVAILABLE = False

try:
    import sympy as sp
    from sympy.printing.mathml import mathml
    SYMPY_AVAILABLE = True
except ImportError:
    SYMPY_AVAILABLE = False

class AdvancedLatexProcessor:
    """Advanced LaTeX to MathML processor with multiple conversion methods"""
    
    def __init__(self):
        self.temp_dir = tempfile.mkdtemp()
        
    def convert_latex_to_mathml(self, latex_string, method='auto'):
        """
        Convert LaTeX to MathML using various methods
        
        Args:
            latex_string (str): LaTeX formula without $ delimiters
            method (str): 'auto', 'latex2mathml', 'sympy', 'pandoc', 'simple'
        
        Returns:
            str: MathML string or formatted text
        """
        # Clean the input
        latex_clean = latex_string.strip()
        
        if method == 'auto':
            # Try methods in order of preference
            for m in ['latex2mathml', 'sympy', 'pandoc', 'simple']:
                try:
                    result = self.convert_latex_to_mathml(latex_clean, m)
                    if result and result != latex_clean:
                        return result
                except:
                    continue
            return self._simple_conversion(latex_clean)
        
        elif method == 'latex2mathml' and LATEX2MATHML_AVAILABLE:
            return self._latex2mathml_conversion(latex_clean)
            
        elif method == 'sympy' and SYMPY_AVAILABLE:
            return self._sympy_conversion(latex_clean)
            
        elif method == 'pandoc':
            return self._pandoc_conversion(latex_clean)
            
        else:  # method == 'simple'
            return self._simple_conversion(latex_clean)
    
    def _latex2mathml_conversion(self, latex_string):
        """Convert using latex2mathml library"""
        try:
            mathml = convert(latex_string)
            return mathml
        except Exception as e:
            print(f"latex2mathml error: {e}")
            return None
    
    def _sympy_conversion(self, latex_string):
        """Convert using SymPy"""
        try:
            # Parse LaTeX using SymPy
            expr = sp.sympify(latex_string.replace('\\', ''))
            mathml_output = mathml(expr)
            return mathml_output
        except Exception as e:
            print(f"SymPy error: {e}")
            return None
    
    def _pandoc_conversion(self, latex_string):
        """Convert using Pandoc"""
        try:
            # Create temporary LaTeX file
            tex_content = f"${latex_string}$"
            tex_file = os.path.join(self.temp_dir, 'temp.tex')
            
            with open(tex_file, 'w', encoding='utf-8') as f:
                f.write(tex_content)
            
            # Use pandoc to convert
            result = subprocess.run([
                'pandoc', '-f', 'latex', '-t', 'mathml', tex_file
            ], capture_output=True, text=True)
            
            if result.returncode == 0:
                return result.stdout.strip()
            else:
                print(f"Pandoc error: {result.stderr}")
                return None
                
        except Exception as e:
            print(f"Pandoc conversion error: {e}")
            return None
    
    def _simple_conversion(self, latex_string):
        """Simple text-based conversion for common LaTeX symbols"""
        
        # Greek letters
        greek_replacements = {
            r'\\alpha': 'α', r'\\beta': 'β', r'\\gamma': 'γ', r'\\delta': 'δ',
            r'\\epsilon': 'ε', r'\\zeta': 'ζ', r'\\eta': 'η', r'\\theta': 'θ',
            r'\\iota': 'ι', r'\\kappa': 'κ', r'\\lambda': 'λ', r'\\mu': 'μ',
            r'\\nu': 'ν', r'\\xi': 'ξ', r'\\pi': 'π', r'\\rho': 'ρ',
            r'\\sigma': 'σ', r'\\tau': 'τ', r'\\upsilon': 'υ', r'\\phi': 'φ',
            r'\\chi': 'χ', r'\\psi': 'ψ', r'\\omega': 'ω',
            r'\\Gamma': 'Γ', r'\\Delta': 'Δ', r'\\Theta': 'Θ', r'\\Lambda': 'Λ',
            r'\\Xi': 'Ξ', r'\\Pi': 'Π', r'\\Sigma': 'Σ', r'\\Phi': 'Φ',
            r'\\Psi': 'Ψ', r'\\Omega': 'Ω'
        }
        
        # Mathematical symbols
        math_replacements = {
            r'\\infty': '∞', r'\\partial': '∂', r'\\nabla': '∇',
            r'\\sum': '∑', r'\\prod': '∏', r'\\int': '∫',
            r'\\oint': '∮', r'\\iint': '∬', r'\\iiint': '∭',
            r'\\pm': '±', r'\\mp': '∓', r'\\times': '×', r'\\div': '÷',
            r'\\cdot': '·', r'\\bullet': '•', r'\\cap': '∩', r'\\cup': '∪',
            r'\\subset': '⊂', r'\\supset': '⊃', r'\\in': '∈', r'\\notin': '∉',
            r'\\exists': '∃', r'\\forall': '∀', r'\\emptyset': '∅',
            r'\\leq': '≤', r'\\geq': '≥', r'\\neq': '≠', r'\\approx': '≈',
            r'\\equiv': '≡', r'\\sim': '∼', r'\\propto': '∝',
            r'\\leftarrow': '←', r'\\rightarrow': '→', r'\\leftrightarrow': '↔',
            r'\\Leftarrow': '⇐', r'\\Rightarrow': '⇒', r'\\Leftrightarrow': '⇔'
        }
        
        # Functions
        function_replacements = {
            r'\\sin': 'sin', r'\\cos': 'cos', r'\\tan': 'tan',
            r'\\sec': 'sec', r'\\csc': 'csc', r'\\cot': 'cot',
            r'\\sinh': 'sinh', r'\\cosh': 'cosh', r'\\tanh': 'tanh',
            r'\\log': 'log', r'\\ln': 'ln', r'\\exp': 'exp',
            r'\\arcsin': 'arcsin', r'\\arccos': 'arccos', r'\\arctan': 'arctan',
            r'\\min': 'min', r'\\max': 'max', r'\\lim': 'lim'
        }
        
        result = latex_string
        
        # Apply all replacements
        all_replacements = {**greek_replacements, **math_replacements, **function_replacements}
        for pattern, replacement in all_replacements.items():
            result = re.sub(pattern, replacement, result)
        
        # Handle superscripts and subscripts
        result = re.sub(r'\^{([^}]+)}', r'^(\1)', result)
        result = re.sub(r'\^(\w)', r'^(\1)', result)
        result = re.sub(r'_{([^}]+)}', r'_(\1)', result)
        result = re.sub(r'_(\w)', r'_(\1)', result)
        
        # Handle fractions
        result = re.sub(r'\\frac{([^}]+)}{([^}]+)}', r'(\1)/(\2)', result)
        
        # Handle square roots
        result = re.sub(r'\\sqrt{([^}]+)}', r'√(\1)', result)
        result = re.sub(r'\\sqrt', '√', result)
        
        # Handle common brackets
        result = result.replace(r'\left(', '(')
        result = result.replace(r'\right)', ')')
        result = result.replace(r'\left[', '[')
        result = result.replace(r'\right]', ']')
        result = result.replace(r'\left\{', '{')
        result = result.replace(r'\right\}', '}')
        
        # Clean up extra backslashes
        result = re.sub(r'\\([a-zA-Z]+)', r'\1', result)
        
        return result
    
    def create_mathml_element(self, mathml_string):
        """Create a Word-compatible MathML element"""
        # This would be used with python-docx to insert actual MathML
        # Implementation depends on the specific Word document structure needed
        
        mathml_template = f"""
        <m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:oMath>
                {mathml_string}
            </m:oMath>
        </m:oMathPara>
        """
        return mathml_template
    
    def batch_convert_formulas(self, text, method='auto'):
        """Convert all LaTeX formulas in text to MathML"""
        
        # Find all formulas
        patterns = [
            (r'\$\{([^}]+)\}\$', r'${{{}\}}$'),  # ${...}$
            (r'\$([^$]+)\$', r'${}$')             # $...$
        ]
        
        result_text = text
        conversions = []
        
        for pattern, replacement_template in patterns:
            matches = list(re.finditer(pattern, result_text))
            
            # Process matches from end to start to preserve positions
            for match in reversed(matches):
                latex_content = match.group(1)
                mathml = self.convert_latex_to_mathml(latex_content, method)
                
                conversions.append({
                    'original': match.group(0),
                    'latex': latex_content,
                    'mathml': mathml,
                    'position': match.span()
                })
                
                # Replace in text
                if mathml:
                    result_text = result_text[:match.start()] + f"[MATH: {mathml}]" + result_text[match.end():]
        
        return result_text, conversions
    
    def cleanup(self):
        """Clean up temporary files"""
        import shutil
        try:
            shutil.rmtree(self.temp_dir)
        except:
            pass

# Example usage and testing
def test_latex_processor():
    """Test the LaTeX processor with common formulas"""
    processor = AdvancedLatexProcessor()
    
    test_formulas = [
        r'\pi r^2',
        r'E = mc^2', 
        r'\frac{a}{b}',
        r'\sqrt{x^2 + y^2}',
        r'\sum_{i=1}^{n} x_i',
        r'\int_0^\infty e^{-x} dx',
        r'\alpha + \beta = \gamma'
    ]
    
    print("Testing LaTeX to MathML conversion:")
    print("="*50)
    
    for formula in test_formulas:
        print(f"LaTeX: ${formula}$")
        
        for method in ['simple', 'latex2mathml', 'sympy']:
            try:
                result = processor.convert_latex_to_mathml(formula, method)
                print(f"  {method:12}: {result}")
            except Exception as e:
                print(f"  {method:12}: Error - {e}")
        
        print("-" * 30)
    
    processor.cleanup()

if __name__ == "__main__":
    test_latex_processor()
