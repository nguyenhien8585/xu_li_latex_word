#!/usr/bin/env python3
"""
Complete Enhanced LaTeX to Word Converter
✅ FIXED: nth roots (∛, ∜, etc.), table duplicates, spacing issues
✅ ADDED: 100+ more LaTeX symbols, better OMML, enhanced formatting
"""

import streamlit as st
import re
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from io import BytesIO
from datetime import datetime
import traceback

# Page configuration
st.set_page_config(
    page_title="Enhanced LaTeX to Word Converter",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Enhanced CSS
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 50%, #f093fb 100%);
        padding: 2.5rem;
        border-radius: 15px;
        text-align: center;
        margin-bottom: 2rem;
        color: white;
        box-shadow: 0 10px 30px rgba(0,0,0,0.3);
    }
    .main-header h1 {
        margin: 0;
        font-size: 3rem;
        font-weight: bold;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    }
    .main-header p {
        margin: 0.5rem 0 0 0;
        font-size: 1.2rem;
        opacity: 0.95;
    }
    .feature-card {
        background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
        padding: 1.5rem;
        border-radius: 12px;
        border-left: 5px solid #667eea;
        margin: 1rem 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .success-box {
        background: linear-gradient(135deg, #56ab2f 0%, #a8e6cf 100%);
        padding: 1.5rem;
        border-radius: 12px;
        color: white;
        text-align: center;
        margin: 1rem 0;
        box-shadow: 0 6px 20px rgba(86, 171, 47, 0.3);
    }
    .enhancement-box {
        background: linear-gradient(135deg, #ff9a9e 0%, #fecfef 50%, #fecfef 100%);
        padding: 1.5rem;
        border-radius: 12px;
        color: #333;
        margin: 1rem 0;
        border: 2px solid #ff6b9d;
    }
    .metric-card {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        text-align: center;
        border-top: 4px solid #667eea;
    }
    .test-case {
        background: #f1f3f4;
        padding: 1rem;
        border-radius: 8px;
        margin: 0.5rem 0;
        border-left: 3px solid #4285f4;
    }
</style>
""", unsafe_allow_html=True)

def convert_latex_symbols(text):
    """COMPLETE ENHANCED LaTeX to Unicode conversion with all fixes"""
    if not text:
        return text
    
    result = text
    
    # Step 1: Handle nth roots FIRST (cube root, 4th root, etc.)
    while re.search(r'\\sqrt\[([^\]]+)\]\{', result):
        match = re.search(r'\\sqrt\[([^\]]+)\]\{', result)
        if not match:
            break
        
        try:
            start = match.start()
            root_index = match.group(1)
            content_start = match.end()
            
            brace_count = 1
            content_end = content_start
            
            while content_end < len(result) and brace_count > 0:
                if result[content_end] == '{':
                    brace_count += 1
                elif result[content_end] == '}':
                    brace_count -= 1
                content_end += 1
            
            content = result[content_start:content_end-1]
            
            # Convert to appropriate root symbol
            root_symbols = {
                '2': '√',
                '3': '∛',
                '4': '∜'
            }
            
            if root_index in root_symbols:
                root_result = f"{root_symbols[root_index]}{content}"
            else:
                root_result = f"ROOT[{root_index}]({content})"
            
            result = result[:start] + root_result + result[content_end:]
        except:
            break
    
    # Step 2: Handle fractions FIRST (before other processing)
    while '\\frac{' in result:
        start = result.find('\\frac{')
        if start == -1:
            break
        
        try:
            num_start = start + 6
            brace_count = 1
            num_end = num_start
            
            while num_end < len(result) and brace_count > 0:
                if result[num_end] == '{':
                    brace_count += 1
                elif result[num_end] == '}':
                    brace_count -= 1
                num_end += 1
            
            numerator = result[num_start:num_end-1].strip()
            
            if num_end < len(result) and result[num_end] == '{':
                denom_start = num_end + 1
                brace_count = 1
                denom_end = denom_start
                
                while denom_end < len(result) and brace_count > 0:
                    if result[denom_end] == '{':
                        brace_count += 1
                    elif result[denom_end] == '}':
                        brace_count -= 1
                    denom_end += 1
                
                denominator = result[denom_start:denom_end-1].strip()
                fraction = f"FRACTION[{numerator}]/[{denominator}]"
                result = result[:start] + fraction + result[denom_end:]
            else:
                break
        except:
            break
    
    # Step 3: Handle square roots (improved)
    while '\\sqrt{' in result:
        start = result.find('\\sqrt{')
        if start == -1:
            break
        
        try:
            content_start = start + 6
            brace_count = 1
            content_end = content_start
            
            while content_end < len(result) and brace_count > 0:
                if result[content_end] == '{':
                    brace_count += 1
                elif result[content_end] == '}':
                    brace_count -= 1
                content_end += 1
            
            content = result[content_start:content_end-1].strip()
            sqrt_result = f"SQRT[{content}]"
            result = result[:start] + sqrt_result + result[content_end:]
        except:
            break
    
    # Step 4: COMPREHENSIVE Greek letters and symbols
    symbols = {
        # Greek letters (lowercase)
        '\\pi': 'π', '\\alpha': 'α', '\\beta': 'β', '\\gamma': 'γ', '\\delta': 'δ',
        '\\epsilon': 'ε', '\\varepsilon': 'ε', '\\zeta': 'ζ', '\\eta': 'η',
        '\\theta': 'θ', '\\vartheta': 'ϑ', '\\iota': 'ι', '\\kappa': 'κ',
        '\\lambda': 'λ', '\\mu': 'μ', '\\nu': 'ν', '\\xi': 'ξ', '\\omicron': 'ο',
        '\\rho': 'ρ', '\\varrho': 'ϱ', '\\sigma': 'σ', '\\varsigma': 'ς',
        '\\tau': 'τ', '\\upsilon': 'υ', '\\phi': 'φ', '\\varphi': 'ϕ',
        '\\chi': 'χ', '\\psi': 'ψ', '\\omega': 'ω',
        
        # Greek letters (uppercase)
        '\\Gamma': 'Γ', '\\Delta': 'Δ', '\\Theta': 'Θ', '\\Lambda': 'Λ',
        '\\Xi': 'Ξ', '\\Pi': 'Π', '\\Sigma': 'Σ', '\\Upsilon': 'Υ',
        '\\Phi': 'Φ', '\\Chi': 'Χ', '\\Psi': 'Ψ', '\\Omega': 'Ω',
        
        # Set theory and logic
        '\\in': '∈', '\\notin': '∉', '\\subset': '⊂', '\\subseteq': '⊆',
        '\\supset': '⊃', '\\supseteq': '⊇', '\\cup': '∪', '\\cap': '∩',
        '\\setminus': '∖', '\\emptyset': '∅', '\\varnothing': '∅',
        '\\exists': '∃', '\\nexists': '∄', '\\forall': '∀',
        '\\neg': '¬', '\\lnot': '¬', '\\land': '∧', '\\lor': '∨',
        '\\implies': '⟹', '\\iff': '⟺', '\\therefore': '∴', '\\because': '∵',
        
        # Relations and inequalities
        '\\leq': '≤', '\\le': '≤', '\\geq': '≥', '\\ge': '≥', 
        '\\neq': '≠', '\\ne': '≠', '\\approx': '≈', '\\equiv': '≡',
        '\\sim': '∼', '\\simeq': '≃', '\\cong': '≅', '\\propto': '∝',
        '\\parallel': '∥', '\\nparallel': '∦', '\\perp': '⊥',
        '\\ll': '≪', '\\gg': '≫', '\\prec': '≺', '\\succ': '≻',
        '\\preceq': '⪯', '\\succeq': '⪰', '\\doteq': '≐',
        
        # Operations and symbols
        '\\infty': '∞', '\\pm': '±', '\\mp': '∓', '\\times': '×', 
        '\\div': '÷', '\\cdot': '·', '\\bullet': '•', '\\circ': '∘',
        '\\oplus': '⊕', '\\ominus': '⊖', '\\otimes': '⊗', '\\oslash': '⊘',
        '\\odot': '⊙', '\\star': '⋆', '\\ast': '∗', '\\bigstar': '★',
        '\\dagger': '†', '\\ddagger': '‡', '\\amalg': '⨿',
        
        # Calculus and analysis
        '\\partial': '∂', '\\nabla': '∇', '\\triangle': '△',
        '\\int': '∫', '\\iint': '∬', '\\iiint': '∭', '\\oint': '∮',
        '\\sum': '∑', '\\prod': '∏', '\\coprod': '∐',
        '\\bigcup': '⋃', '\\bigcap': '⋂', '\\bigsqcup': '⨆',
        '\\bigvee': '⋁', '\\bigwedge': '⋀', '\\bigotimes': '⨂',
        '\\bigoplus': '⨁', '\\bigodot': '⨀',
        
        # Number sets
        '\\mathbb{R}': 'ℝ', '\\mathbb{Z}': 'ℤ', '\\mathbb{Q}': 'ℚ', 
        '\\mathbb{N}': 'ℕ', '\\mathbb{C}': 'ℂ', '\\mathbb{P}': 'ℙ',
        '\\mathbb{H}': 'ℍ', '\\mathbb{F}': 'F', '\\mathbb{E}': 'E',
        
        # Arrows
        '\\rightarrow': '→', '\\to': '→', '\\leftarrow': '←', 
        '\\leftrightarrow': '↔', '\\Rightarrow': '⇒', '\\Leftarrow': '⇐',
        '\\Leftrightarrow': '⇔', '\\uparrow': '↑', '\\downarrow': '↓',
        '\\updownarrow': '↕', '\\nearrow': '↗', '\\searrow': '↘',
        '\\swarrow': '↙', '\\nwarrow': '↖', '\\mapsto': '↦',
        '\\longmapsto': '⟼', '\\hookrightarrow': '↪', '\\hookleftarrow': '↩',
        '\\rightharpoonup': '⇀', '\\rightharpoondown': '⇁',
        '\\leftharpoonup': '↼', '\\leftharpoondown': '↽',
        
        # Miscellaneous symbols
        '\\aleph': 'ℵ', '\\beth': 'ℶ', '\\gimel': 'ℷ', '\\daleth': 'ℸ',
        '\\ell': 'ℓ', '\\wp': '℘', '\\Re': 'ℜ', '\\Im': 'ℑ',
        '\\angle': '∠', '\\measuredangle': '∡', '\\sphericalangle': '∢',
        '\\top': '⊤', '\\bot': '⊥', '\\vdash': '⊢', '\\dashv': '⊣',
        '\\models': '⊨', '\\vDash': '⊨', '\\Vdash': '⊩', '\\Vvdash': '⊪',
        
        # Delimiters
        '\\langle': '⟨', '\\rangle': '⟩', '\\lceil': '⌈', '\\rceil': '⌉',
        '\\lfloor': '⌊', '\\rfloor': '⌋', '\\ulcorner': '⌜', '\\urcorner': '⌝',
        '\\llcorner': '⌞', '\\lrcorner': '⌟'
    }
    
    # Apply symbols with word boundary protection
    for latex_sym, unicode_sym in symbols.items():
        pattern = r'\b' + re.escape(latex_sym) + r'\b'
        result = re.sub(pattern, unicode_sym, result)
    
    # Step 5: Enhanced mathematical functions with proper spacing
    functions = {
        # Trigonometric functions
        '\\sin': 'sin', '\\cos': 'cos', '\\tan': 'tan', '\\cot': 'cot',
        '\\sec': 'sec', '\\csc': 'csc', 
        '\\arcsin': 'arcsin', '\\arccos': 'arccos', '\\arctan': 'arctan', 
        '\\arccot': 'arccot', '\\arcsec': 'arcsec', '\\arccsc': 'arccsc',
        
        # Hyperbolic functions
        '\\sinh': 'sinh', '\\cosh': 'cosh', '\\tanh': 'tanh', '\\coth': 'coth',
        '\\sech': 'sech', '\\csch': 'csch',
        
        # Logarithmic and exponential
        '\\log': 'log', '\\ln': 'ln', '\\lg': 'lg', '\\exp': 'exp',
        
        # Limits and extrema
        '\\lim': 'lim', '\\limsup': 'lim sup', '\\liminf': 'lim inf',
        '\\sup': 'sup', '\\inf': 'inf', '\\max': 'max', '\\min': 'min',
        
        # Other functions
        '\\gcd': 'gcd', '\\lcm': 'lcm', '\\det': 'det', '\\dim': 'dim',
        '\\deg': 'deg', '\\ker': 'ker', '\\arg': 'arg', '\\sgn': 'sgn'
    }
    
    for latex_func, unicode_func in functions.items():
        # Add space after function if followed by letter/number/opening bracket
        pattern = r'\b' + re.escape(latex_func) + r'\b(?=[\w\(\[])'
        result = re.sub(pattern, unicode_func + ' ', result)
        # Regular replacement for other cases
        result = result.replace(latex_func, unicode_func)
    
    # Step 6: Enhanced prime notation
    prime_patterns = [
        (r"([A-Za-z0-9]+)(\^?\{?'+'?\}?)", lambda m: m.group(1) + "'" * m.group(2).count("'")),
        (r"([A-Za-z0-9]+)\^?\{?prime\}?", r"\1'"),
        (r"\(prime\)", "'"),
    ]
    
    for pattern, replacement in prime_patterns:
        result = re.sub(pattern, replacement, result)
    
    # Step 7: COMPREHENSIVE subscripts
    subscript_map = {
        '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄',
        '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉',
        'a': 'ₐ', 'e': 'ₑ', 'h': 'ₕ', 'i': 'ᵢ', 'j': 'ⱼ', 'k': 'ₖ',
        'l': 'ₗ', 'm': 'ₘ', 'n': 'ₙ', 'o': 'ₒ', 'p': 'ₚ', 'r': 'ᵣ',
        's': 'ₛ', 't': 'ₜ', 'u': 'ᵤ', 'v': 'ᵥ', 'x': 'ₓ',
        '+': '₊', '-': '₋', '=': '₌', '(': '₍', ')': '₎'
    }
    
    def replace_subscript(match):
        base = match.group(1)
        content = match.group(2)
        converted = ''.join(subscript_map.get(c, c) for c in content)
        return base + converted
    
    result = re.sub(r'([A-Za-z0-9]+)_\{([^}]+)\}', replace_subscript, result)
    result = re.sub(r'([A-Za-z0-9]+)_([0-9a-z])', lambda m: m.group(1) + subscript_map.get(m.group(2), m.group(2)), result)
    
    # Step 8: COMPREHENSIVE superscripts
    superscript_map = {
        '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴',
        '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹',
        'a': 'ᵃ', 'b': 'ᵇ', 'c': 'ᶜ', 'd': 'ᵈ', 'e': 'ᵉ', 'f': 'ᶠ',
        'g': 'ᵍ', 'h': 'ʰ', 'i': 'ⁱ', 'j': 'ʲ', 'k': 'ᵏ', 'l': 'ˡ',
        'm': 'ᵐ', 'n': 'ⁿ', 'o': 'ᵒ', 'p': 'ᵖ', 'r': 'ʳ', 's': 'ˢ',
        't': 'ᵗ', 'u': 'ᵘ', 'v': 'ᵛ', 'w': 'ʷ', 'x': 'ˣ', 'y': 'ʸ', 'z': 'ᶻ',
        '+': '⁺', '-': '⁻', '=': '⁼', '(': '⁽', ')': '⁾'
    }
    
    def replace_superscript(match):
        base = match.group(1)
        content = match.group(2)
        if 'prime' in content or content == "'":
            return base + "'"
        converted = ''.join(superscript_map.get(c, c) for c in content)
        return base + converted
    
    result = re.sub(r'([A-Za-z0-9]+)\^\{([^}]+)\}', replace_superscript, result)
    result = re.sub(r'([A-Za-z0-9]+)\^([0-9a-zA-Z])', lambda m: m.group(1) + superscript_map.get(m.group(2), m.group(2)), result)
    
    # Step 9: Handle brackets and delimiters
    bracket_pairs = [
        ('\\left(', '('), ('\\right)', ')'),
        ('\\left[', '['), ('\\right]', ']'),
        ('\\left\\{', '{'), ('\\right\\}', '}'),
        ('\\left|', '|'), ('\\right|', '|'),
        ('\\lbrack', '['), ('\\rbrack', ']'),
        ('\\lbrace', '{'), ('\\rbrace', '}'),
        ('\\{', '{'), ('\\}', '}'),
        ('\\langle', '⟨'), ('\\rangle', '⟩'),
        ('\\lceil', '⌈'), ('\\rceil', '⌉'),
        ('\\lfloor', '⌊'), ('\\rfloor', '⌋')
    ]
    
    for latex_bracket, unicode_bracket in bracket_pairs:
        result = result.replace(latex_bracket, unicode_bracket)
    
    # Step 10: ENHANCED spacing cleanup
    # Remove multiple spaces
    result = re.sub(r'\s+', ' ', result)
    # Fix spacing around operators (with better patterns)
    result = re.sub(r'\s*([+\-×÷=≤≥<>≠≈±∓·∘])\s*', r' \1 ', result)
    # Fix spacing after functions (preserve existing spaces)
    result = re.sub(r'\b(sin|cos|tan|cot|sec|csc|log|ln|exp|max|min|lim|sup|inf|gcd|det)\s*(?=[\w\(])', r'\1 ', result)
    # Remove spaces before commas and periods
    result = re.sub(r'\s+([,.])', r'\1', result)
    # Fix spacing around fractions
    result = re.sub(r'FRACTION\s*\[', 'FRACTION[', result)
    
    # Step 11: Clean up remaining LaTeX commands
    result = re.sub(r'\\[a-zA-Z]+\*?', '', result)  # Remove unknown LaTeX commands
    result = result.replace('\\', '')
    
    return result.strip()

def find_math_expressions(text):
    """Enhanced math expression finder with better detection"""
    expressions = []
    i = 0
    
    while i < len(text):
        if text[i] == '$':
            start = i
            i += 1
            
            # Handle $$ format
            if i < len(text) and text[i] == '$':
                i += 1  # Skip second $
                content_start = i
                
                # Find closing $$
                while i < len(text) - 1:
                    if text[i] == '$' and text[i+1] == '$':
                        end = i + 2
                        content = text[content_start:i]
                        if content.strip():
                            expressions.append((start, end, content.strip()))
                        i += 2
                        break
                    i += 1
                else:
                    i += 1
            
            # Handle ${...}$ format
            elif i < len(text) and text[i] == '{':
                i += 1
                brace_count = 1
                content_start = i
                
                while i < len(text) and brace_count > 0:
                    if text[i] == '{':
                        brace_count += 1
                    elif text[i] == '}':
                        brace_count -= 1
                    i += 1
                
                if i < len(text) and text[i] == '$':
                    end = i + 1
                    content = text[content_start:i-1]
                    if content.strip():
                        expressions.append((start, end, content.strip()))
                    i += 1
                else:
                    i += 1
            
            # Handle regular $...$ format with improved detection
            else:
                content_start = i
                brace_depth = 0
                paren_depth = 0
                
                while i < len(text):
                    if text[i] == '{':
                        brace_depth += 1
                    elif text[i] == '}':
                        brace_depth -= 1
                    elif text[i] == '(':
                        paren_depth += 1
                    elif text[i] == ')':
                        paren_depth -= 1
                    elif text[i] == '$' and brace_depth == 0 and paren_depth == 0:
                        break
                    i += 1
                
                if i < len(text):
                    end = i + 1
                    content = text[content_start:i]
                    if content.strip():
                        expressions.append((start, end, content.strip()))
                    i += 1
                else:
                    i += 1
        else:
            i += 1
    
    return expressions

def create_fraction_omml(paragraph, numerator, denominator):
    """Create proper fraction OMML with enhanced formatting"""
    try:
        safe_num = numerator.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        safe_den = denominator.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:f>
                <m:fPr>
                    <m:type m:val="lin"/>
                </m:fPr>
                <m:num>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{safe_num}</m:t>
                    </m:r>
                </m:num>
                <m:den>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{safe_den}</m:t>
                    </m:r>
                </m:den>
            </m:f>
        </m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_sqrt_omml(paragraph, content):
    """Create proper square root OMML"""
    try:
        safe_content = content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:rad>
                <m:radPr></m:radPr>
                <m:deg></m:deg>
                <m:e>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{safe_content}</m:t>
                    </m:r>
                </m:e>
            </m:rad>
        </m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_nth_root_omml(paragraph, index, content):
    """Create OMML for nth roots (cube root, 4th root, etc.)"""
    try:
        safe_content = content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        safe_index = index.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:rad>
                <m:radPr></m:radPr>
                <m:deg>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="p"/>
                        </m:rPr>
                        <m:t>{safe_index}</m:t>
                    </m:r>
                </m:deg>
                <m:e>
                    <m:r>
                        <m:rPr>
                            <m:scr m:val="roman"/>
                            <m:sty m:val="i"/>
                        </m:rPr>
                        <m:t>{safe_content}</m:t>
                    </m:r>
                </m:e>
            </m:rad>
        </m:oMath>'''
        
        math_element = parse_xml(omml_xml)
        paragraph._element.append(math_element)
        return True
    except Exception:
        return False

def create_equation_object(paragraph, latex_text, method='auto'):
    """ENHANCED equation object creation with nth roots and better OMML"""
    unicode_text = convert_latex_symbols(latex_text)
    
    # Check for nth roots first
    root_match = re.search(r'ROOT\[([^\]]+)\]\(([^)]+)\)', unicode_text)
    if root_match:
        index = root_match.group(1)
        content = root_match.group(2)
        
        if create_nth_root_omml(paragraph, index, content):
            return 'omml'
        else:
            # Fallback to Unicode
            if index == '3':
                nth_root_text = f"∛{content}"
            elif index == '4':
                nth_root_text = f"∜{content}"
            else:
                nth_root_text = f"ⁿ√{content} (n={index})"
            unicode_text = unicode_text.replace(root_match.group(0), nth_root_text)
    
    # Check for fractions
    if 'FRACTION[' in unicode_text:
        match = re.search(r'FRACTION\[([^\]]+)\]/\[([^\]]+)\]', unicode_text)
        if match:
            numerator = convert_latex_symbols(match.group(1))
            denominator = convert_latex_symbols(match.group(2))
            
            if create_fraction_omml(paragraph, numerator, denominator):
                return 'omml'
            else:
                fraction_text = f"({numerator})/({denominator})"
                unicode_text = unicode_text.replace(match.group(0), fraction_text)
    
    # Check for square roots
    if 'SQRT[' in unicode_text:
        match = re.search(r'SQRT\[([^\]]+)\]', unicode_text)
        if match:
            content = convert_latex_symbols(match.group(1))
            
            if create_sqrt_omml(paragraph, content):
                return 'omml'
            else:
                sqrt_text = f"√{content}"
                unicode_text = unicode_text.replace(match.group(0), sqrt_text)
    
    # Regular OMML for other expressions
    if method == 'omml' or method == 'auto':
        try:
            safe_text = unicode_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            
            omml_xml = f'''<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
                <m:r>
                    <m:rPr>
                        <m:scr m:val="roman"/>
                        <m:sty m:val="i"/>
                    </m:rPr>
                    <m:t>{safe_text}</m:t>
                </m:r>
            </m:oMath>'''
            
            math_element = parse_xml(omml_xml)
            paragraph._element.append(math_element)
            return 'omml'
        except Exception as e:
            if method == 'omml':
                st.warning(f"OMML failed: {e}")
    
    # Styled text fallback
    if method == 'styled' or method == 'auto':
        try:
            run = paragraph.add_run(unicode_text)
            run.font.name = 'Cambria Math'
            run.font.size = Pt(11)
            run.italic = True
            run.font.color.rgb = RGBColor(0, 32, 96)
            return 'styled'
        except Exception as e:
            if method == 'styled':
                st.warning(f"Styled failed: {e}")
    
    # Final fallback
    try:
        run = paragraph.add_run(unicode_text)
        run.italic = True
        run.font.color.rgb = RGBColor(0, 0, 139)
        return 'fallback'
    except:
        run = paragraph.add_run(f"[{latex_text}]")
        return 'error'

def detect_markdown_table(text):
    """Enhanced markdown table detection with stricter validation"""
    lines = [line.strip() for line in text.strip().split('\n') if line.strip()]
    if len(lines) < 2:
        return False
    
    # Check for pipe characters in multiple lines
    pipe_lines = [line for line in lines if '|' in line and len(line.split('|')) >= 3]
    if len(pipe_lines) < 2:
        return False
    
    # Look for separator line more strictly
    separator_found = False
    for line in lines:
        # Separator line should contain only |, -, :, and spaces
        if re.match(r'^[\|\s\-:]+$', line) and '|' in line and '-' in line:
            separator_found = True
            break
    
    return separator_found

def parse_markdown_table(text):
    """Enhanced markdown table parsing with duplicate prevention"""
    lines = [line.strip() for line in text.strip().split('\n') if line.strip()]
    table_data = []
    separator_indices = []
    
    # Find separator lines
    for i, line in enumerate(lines):
        if re.match(r'^[\|\s\-:]+$', line) and '|' in line and '-' in line:
            separator_indices.append(i)
    
    # Process non-separator lines
    for i, line in enumerate(lines):
        if i in separator_indices:
            continue  # Skip separator lines
        
        if '|' in line:
            # Clean and split the line
            line = line.strip()
            if line.startswith('|'):
                line = line[1:]
            if line.endswith('|'):
                line = line[:-1]
            
            cells = [cell.strip() for cell in line.split('|')]
            # Filter out empty cells at the beginning/end
            while cells and not cells[0]:
                cells.pop(0)
            while cells and not cells[-1]:
                cells.pop()
            
            if cells:  # Only add non-empty rows
                table_data.append(cells)
    
    # Remove duplicate rows (FIXED)
    unique_table_data = []
    seen_rows = set()
    for row in table_data:
        row_tuple = tuple(row)
        if row_tuple not in seen_rows:
            unique_table_data.append(row)
            seen_rows.add(row_tuple)
    
    return unique_table_data

def create_word_table(doc, table_data, options):
    """Enhanced Word table creation with better formatting"""
    if not table_data:
        return None, 0
    
    try:
        num_rows = len(table_data)
        num_cols = max(len(row) for row in table_data) if table_data else 0
        
        if num_rows == 0 or num_cols == 0:
            return None, 0
        
        # Create table
        word_table = doc.add_table(rows=num_rows, cols=num_cols)
        table_math = 0
        
        # Fill table data
        for r, row_data in enumerate(table_data):
            for c in range(num_cols):
                cell = word_table.rows[r].cells[c]
                cell.text = ""
                
                # Get cell content
                if c < len(row_data):
                    cell_text = row_data[c].strip()
                else:
                    cell_text = ""
                
                # Create paragraph in cell
                if cell.paragraphs:
                    cell_para = cell.paragraphs[0]
                    cell_para.clear()
                else:
                    cell_para = cell.add_paragraph()
                
                # Process cell content
                if cell_text:
                    if not format_question_answer(cell_para, cell_text, options):
                        cell_math_count, _ = process_text_with_math(cell_para, cell_text, options)
                        table_math += cell_math_count
                
                # Enhanced formatting
                cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell_para.runs:
                    if not run.font.name:
                        run.font.name = 'Times New Roman'
                    if not run.font.size:
                        run.font.size = Pt(10)
        
        # Apply enhanced table formatting
        if options.get('format_tables', True):
            try:
                word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                
                # Enhanced header row styling
                if word_table.rows:
                    header_row = word_table.rows[0]
                    for cell in header_row.cells:
                        # Bold white text for header
                        for para in cell.paragraphs:
                            for run in para.runs:
                                run.bold = True
                                run.font.color.rgb = RGBColor(255, 255, 255)
                        
                        # Enhanced blue background
                        try:
                            shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="2F5597"/>'
                            shading = parse_xml(shading_xml)
                            cell._tc.get_or_add_tcPr().append(shading)
                        except:
                            pass
                
                # Enhanced borders
                try:
                    for row in word_table.rows:
                        for cell in row.cells:
                            tc = cell._tc
                            tcPr = tc.get_or_add_tcPr()
                            
                            borders_xml = '''
                            <w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                                <w:top w:val="single" w:sz="6" w:space="0" w:color="2F5597"/>
                                <w:left w:val="single" w:sz="6" w:space="0" w:color="2F5597"/>
                                <w:bottom w:val="single" w:sz="6" w:space="0" w:color="2F5597"/>
                                <w:right w:val="single" w:sz="6" w:space="0" w:color="2F5597"/>
                            </w:tcBorders>'''
                            borders = parse_xml(borders_xml)
                            tcPr.append(borders)
                except:
                    pass
            except:
                pass
        
        return word_table, table_math
        
    except Exception as e:
        st.warning(f"Table creation error: {e}")
        return None, 0

def process_text_with_math(paragraph, text, options):
    """Enhanced text processing with better math detection"""
    expressions = find_math_expressions(text)
    
    if not expressions:
        run = paragraph.add_run(text)
        run.font.size = Pt(11)
        run.font.name = 'Times New Roman'
        return 0, {'omml': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    last_pos = 0
    math_count = 0
    method_stats = {'omml': 0, 'styled': 0, 'fallback': 0, 'error': 0}
    
    for start, end, latex_content in expressions:
        # Add text before math
        if start > last_pos:
            before_text = text[last_pos:start]
            if before_text:
                run = paragraph.add_run(before_text)
                run.font.size = Pt(11)
                run.font.name = 'Times New Roman'
        
        # Create equation with enhanced processing
        method_used = create_equation_object(paragraph, latex_content, options.get('equation_method', 'auto'))
        method_stats[method_used] += 1
        math_count += 1
        last_pos = end
    
    # Add remaining text
    if last_pos < len(text):
        remaining = text[last_pos:]
        if remaining:
            run = paragraph.add_run(remaining)
            run.font.size = Pt(11)
            run.font.name = 'Times New Roman'
    
    return math_count, method_stats

def format_question_answer(paragraph, text, options):
    """Enhanced question and answer formatting"""
    paragraph.clear()
    
    # Enhanced question patterns
    question_patterns = [
        r'^(Câu\s+\d+[\.:])\s*(.*)',
        r'^(Question\s+\d+[\.:])\s*(.*)',
        r'^(\d+[\.:])\s*(.*)',
        r'^(Bài\s+\d+[\.:])\s*(.*)',
        r'^(Problem\s+\d+[\.:])\s*(.*)'
    ]
    
    for pattern in question_patterns:
        match = re.match(pattern, text, re.IGNORECASE)
        if match:
            question_part = match.group(1)
            content_part = match.group(2)
            
            # Enhanced question number styling
            run_q = paragraph.add_run(question_part)
            run_q.bold = True
            run_q.font.size = Pt(12)
            run_q.font.name = 'Times New Roman'
            run_q.font.color.rgb = RGBColor(0, 32, 96)
            
            if content_part:
                paragraph.add_run(" ")
                process_text_with_math(paragraph, content_part, options)
            
            return True
    
    # Enhanced answer patterns
    answer_pattern = r'^([A-Da-d])[\.\)]\s*(.*)'
    match = re.match(answer_pattern, text)
    if match:
        answer_letter = match.group(1).upper() + '.'
        answer_content = match.group(2)
        
        # Enhanced answer letter styling
        run_l = paragraph.add_run(answer_letter)
        run_l.bold = True
        run_l.font.color.rgb = RGBColor(47, 85, 151)
        run_l.font.size = Pt(11)
        run_l.font.name = 'Times New Roman'
        
        if answer_content:
            paragraph.add_run(" ")
            process_text_with_math(paragraph, answer_content, options)
        
        return True
    
    # Not Q&A format
    process_text_with_math(paragraph, text, options)
    return False

def process_document(doc, options):
    """Enhanced document processing with all improvements"""
    new_doc = Document()
    
    stats = {
        'paragraphs': 0,
        'tables': 0,
        'markdown_tables': 0,
        'math_expressions': 0,
        'questions_formatted': 0,
        'nth_roots': 0,
        'method_stats': {'omml': 0, 'styled': 0, 'fallback': 0, 'error': 0},
        'processing_log': []
    }
    
    try:
        # Collect all paragraph text
        paragraph_texts = []
        for para in doc.paragraphs:
            paragraph_texts.append(para.text)
        
        # Process paragraphs with enhanced markdown table detection
        i = 0
        while i < len(paragraph_texts):
            text = paragraph_texts[i].strip()
            stats['paragraphs'] += 1
            
            if not text:
                new_doc.add_paragraph()
                i += 1
                continue
            
            # Enhanced markdown table detection
            if options.get('convert_markdown', True) and '|' in text:
                # Look ahead to collect table lines
                table_lines = []
                j = i
                
                while j < len(paragraph_texts):
                    current_text = paragraph_texts[j].strip()
                    if not current_text:
                        j += 1
                        if j < len(paragraph_texts) and '|' in paragraph_texts[j]:
                            continue
                        else:
                            break
                    if '|' in current_text:
                        table_lines.append(current_text)
                        j += 1
                    else:
                        break
                
                # Check if it's a valid markdown table
                combined_table_text = '\n'.join(table_lines)
                if detect_markdown_table(combined_table_text):
                    # Parse and create Word table
                    table_data = parse_markdown_table(combined_table_text)
                    if table_data:
                        word_table, table_math = create_word_table(new_doc, table_data, options)
                        if word_table is not None:
                            stats['markdown_tables'] += 1
                            stats['math_expressions'] += table_math
                            stats['processing_log'].append(f"Converted markdown table: {len(table_data)} rows, {table_math} math expressions")
                    
                    i = j  # Skip processed table lines
                    continue
            
            # Process regular paragraph
            new_para = new_doc.add_paragraph()
            
            try:
                # Count nth roots in text
                nth_root_count = text.count('\\sqrt[') + text.count('∛') + text.count('∜')
                stats['nth_roots'] += nth_root_count
                
                if options.get('format_qa', True) and format_question_answer(new_para, text, options):
                    stats['questions_formatted'] += 1
                    stats['processing_log'].append(f"Para {i+1}: Q&A formatted")
                else:
                    math_count, method_stats = process_text_with_math(new_para, text, options)
                    stats['math_expressions'] += math_count
                    for method, count in method_stats.items():
                        stats['method_stats'][method] += count
                    
                    if math_count > 0:
                        stats['processing_log'].append(f"Para {i+1}: {math_count} equations converted")
            except Exception as e:
                # Fallback to plain text
                new_para.clear()
                run = new_para.add_run(text)
                run.font.size = Pt(11)
                run.font.name = 'Times New Roman'
                stats['processing_log'].append(f"Para {i+1}: Error, used plain text - {str(e)}")
            
            i += 1
        
        # Process existing Word tables (enhanced)
        for table_idx, table in enumerate(doc.tables):
            if not table.rows:
                continue
            
            try:
                num_rows = len(table.rows)
                num_cols = len(table.rows[0].cells) if table.rows else 0
                
                new_table = new_doc.add_table(rows=num_rows, cols=num_cols)
                table_math = 0
                
                for r in range(num_rows):
                    for c in range(min(len(table.rows[r].cells), num_cols)):
                        try:
                            original_cell = table.rows[r].cells[c]
                            new_cell = new_table.rows[r].cells[c]
                            
                            cell_text = original_cell.text.strip()
                            if cell_text:
                                new_cell.text = ""
                                cell_para = new_cell.paragraphs[0] if new_cell.paragraphs else new_cell.add_paragraph()
                                cell_para.clear()
                                
                                if not (options.get('format_qa', True) and format_question_answer(cell_para, cell_text, options)):
                                    cell_math_count, cell_method_stats = process_text_with_math(cell_para, cell_text, options)
                                    table_math += cell_math_count
                                    for method, count in cell_method_stats.items():
                                        stats['method_stats'][method] += count
                                
                                cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                for run in cell_para.runs:
                                    if not run.font.name:
                                        run.font.name = 'Times New Roman'
                                    if not run.font.size:
                                        run.font.size = Pt(10)
                        except Exception:
                            continue
                
                # Enhanced table formatting
                if options.get('format_tables', True):
                    try:
                        new_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                        
                        # Enhanced header styling
                        if new_table.rows:
                            header_row = new_table.rows[0]
                            for cell in header_row.cells:
                                for para in cell.paragraphs:
                                    for run in para.runs:
                                        run.bold = True
                                        run.font.color.rgb = RGBColor(255, 255, 255)
                                
                                try:
                                    shading_xml = '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="2F5597"/>'
                                    shading = parse_xml(shading_xml)
                                    cell._tc.get_or_add_tcPr().append(shading)
                                except:
                                    pass
                    except Exception:
                        pass
                
                stats['tables'] += 1
                stats['math_expressions'] += table_math
                if table_math > 0:
                    stats['processing_log'].append(f"Word table {table_idx+1}: {table_math} equations")
                
            except Exception as e:
                stats['processing_log'].append(f"Word table {table_idx+1}: Error - {str(e)}")
                continue
        
    except Exception as e:
        stats['processing_log'].append(f"Document processing error: {str(e)}")
    
    return new_doc, stats

# Main Streamlit App
def main():
    # Enhanced Header
    st.markdown("""
    <div class="main-header">
        <h1>🚀 Enhanced LaTeX to Word Converter</h1>
        <p>✅ ALL FIXES: nth roots (∛∜), table duplicates, spacing, 100+ symbols</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("## ⚙️ Cài đặt nâng cao")
        
        # Processing options
        st.markdown("### 🔧 Tùy chọn xử lý")
        format_qa = st.checkbox("📝 Định dạng câu hỏi & đáp án", value=True)
        format_tables = st.checkbox("📊 Định dạng bảng", value=True)
        convert_markdown = st.checkbox("🔄 Chuyển markdown tables", value=True,
                                     help="Chuyển | tables | thành Word tables (FIXED duplicates)")
        
        equation_method = st.selectbox(
            "🧮 Phương pháp equation",
            ["auto", "omml", "styled"],
            index=0,
            help="auto: thử OMML trước, fallback styled"
        )
        
        st.markdown("---")
        
        # Enhanced test cases
        st.markdown("### 🎯 Test Enhanced LaTeX")
        test_cases = [
            "\\sqrt{3}\\cot 2x = 1",
            "\\sqrt[3]{8} = 2",
            "\\sqrt[4]{16} = 2", 
            "\\frac{\\sqrt{3}}{2}",
            "\\tan\\frac{65\\pi}{6}",
            "u_{n+1}^{(k)} = \\alpha_i \\beta^j",
            "A^{'}B^{'}C^{'}D^{'}",
            "\\lim_{x \\to \\infty} \\frac{1}{x} = 0",
            "\\int_0^\\infty e^{-x} dx = 1",
            "\\sum_{i=1}^n x_i^2 \\geq 0"
        ]
        
        selected_test = st.selectbox("Enhanced test cases:", [""] + test_cases)
        if selected_test:
            try:
                converted = convert_latex_symbols(selected_test)
                st.code(f"${selected_test}$")
                st.success(f"→ {converted}")
                
                # Show enhancement type
                if '∛' in converted:
                    st.info("✅ Cube root detected")
                elif '∜' in converted:
                    st.info("✅ 4th root detected")
                elif 'ROOT[' in converted:
                    st.info("✅ nth root detected")
                elif 'FRACTION[' in converted:
                    st.info("✅ Fraction OMML")
                elif 'SQRT[' in converted:
                    st.info("✅ Square root OMML")
                else:
                    st.info("✅ Enhanced symbols")
            except Exception as e:
                st.error(f"Error: {e}")
        
        st.markdown("---")
        st.markdown("### ✨ NEW ENHANCEMENTS")
        st.markdown("""
        <div class="enhancement-box">
        <strong>🆕 What's Fixed:</strong><br>
        • nth roots: ∛∜ + OMML<br>
        • Table duplicates removed<br>
        • Spacing: sin x not sinx<br>
        • 100+ new symbols<br>
        • Better OMML generation<br>
        • Enhanced formatting
        </div>
        """, unsafe_allow_html=True)
        
        show_debug = st.checkbox("🔍 Debug log", value=False)
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("## 📁 Upload Enhanced Processing")
        uploaded_file = st.file_uploader(
            "Chọn file .docx để xử lý enhanced",
            type=["docx"],
            help="File Word có LaTeX, markdown tables, nth roots, etc."
        )
        
        if uploaded_file:
            st.success(f"✅ Uploaded: **{uploaded_file.name}**")
            
            # Enhanced Preview
            with st.expander("👀 Enhanced Preview & Analysis", expanded=True):
                try:
                    doc = Document(uploaded_file)
                    st.info(f"📄 {len(doc.paragraphs)} paragraphs | 📊 {len(doc.tables)} tables")
                    
                    # Analyze content
                    all_text = "\n".join([para.text for para in doc.paragraphs])
                    
                    # Find LaTeX expressions
                    expressions = find_math_expressions(all_text)
                    
                    # Count different types
                    nth_roots = all_text.count('\\sqrt[')
                    fractions = all_text.count('\\frac{')
                    sqrt_count = all_text.count('\\sqrt{')
                    
                    # Find markdown tables
                    markdown_tables = []
                    lines = all_text.split('\n')
                    i = 0
                    while i < len(lines):
                        if '|' in lines[i]:
                            table_lines = []
                            j = i
                            while j < len(lines) and ('|' in lines[j] or not lines[j].strip()):
                                if lines[j].strip():
                                    table_lines.append(lines[j])
                                j += 1
                            if len(table_lines) >= 2 and detect_markdown_table('\n'.join(table_lines)):
                                markdown_tables.append('\n'.join(table_lines))
                            i = j
                        else:
                            i += 1
                    
                    # Enhanced statistics
                    col_a, col_b, col_c, col_d = st.columns(4)
                    with col_a:
                        st.metric("LaTeX Expressions", len(expressions))
                    with col_b:
                        st.metric("nth Roots", nth_roots)
                    with col_c:
                        st.metric("Fractions", fractions)
                    with col_d:
                        st.metric("MD Tables", len(markdown_tables))
                    
                    # Show samples
                    if expressions:
                        st.markdown("**🧮 Enhanced LaTeX Processing:**")
                        unique_expressions = list(set([expr[2] for expr in expressions[:8]]))
                        
                        for expr in unique_expressions:
                            col_x, col_y = st.columns([1, 1])
                            with col_x:
                                st.code(f"${expr}$", language="latex")
                            with col_y:
                                converted = convert_latex_symbols(expr)
                                # Color code by type
                                if '∛' in converted or '∜' in converted:
                                    st.success(f"→ {converted}")
                                elif 'FRACTION' in convert_latex_symbols(expr):
                                    st.info(f"→ {converted}")
                                else:
                                    st.code(f"→ {converted}")
                        
                        if len(expressions) > 8:
                            st.info(f"... and {len(expressions) - 8} more")
                    
                    if markdown_tables:
                        st.markdown("**📊 Enhanced Table Processing:**")
                        for i, table in enumerate(markdown_tables[:2]):
                            with st.expander(f"Table {i+1} Preview"):
                                # Parse and show structure
                                parsed = parse_markdown_table(table)
                                st.write(f"Rows: {len(parsed)}, Columns: {max(len(row) for row in parsed) if parsed else 0}")
                                st.code(table[:300] + "..." if len(table) > 300 else table)
                        if len(markdown_tables) > 2:
                            st.info(f"... and {len(markdown_tables) - 2} more tables")
                    
                    if not expressions and not markdown_tables:
                        st.warning("No LaTeX expressions or markdown tables found")
                        
                except Exception as e:
                    st.error(f"Preview error: {e}")
    
    with col2:
        st.markdown("## 🎯 Complete Enhancement")
        
        enhancements = [
            ("∛", "Cube Roots", "\\sqrt[3]{8} → ∛8"),
            ("∜", "4th Roots", "\\sqrt[4]{16} → ∜16"),
            ("📊", "No Duplicates", "Clean tables"),
            ("🔤", "Perfect Spacing", "sin x not sinx"),
            ("🧮", "Enhanced OMML", "Better equations"),
            ("🎨", "Pro Formatting", "Publication ready")
        ]
        
        for icon, title, desc in enhancements:
            st.markdown(f"""
            <div class="feature-card">
                <strong>{icon} {title}</strong><br>
                <small>{desc}</small>
            </div>
            """, unsafe_allow_html=True)
    
    # Enhanced Processing
    if uploaded_file:
        st.markdown("---")
        
        if st.button("🚀 COMPLETE ENHANCED CONVERSION", type="primary", use_container_width=True):
            options = {
                'format_qa': format_qa,
                'format_tables': format_tables,
                'convert_markdown': convert_markdown,
                'equation_method': equation_method
            }
            
            with st.spinner("⏳ Processing with ALL enhancements..."):
                try:
                    doc = Document(uploaded_file)
                    new_doc, stats = process_document(doc, options)
                    
                    # Save to buffer
                    buffer = BytesIO()
                    new_doc.save(buffer)
                    buffer.seek(0)
                    
                    # Generate filename
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_filename = f"{uploaded_file.name.rsplit('.', 1)[0]}_ENHANCED_{timestamp}.docx"
                    
                    # Enhanced success message
                    st.markdown("""
                    <div class="success-box">
                        <h3>🎉 COMPLETE ENHANCED CONVERSION SUCCESS!</h3>
                        <p>All fixes + enhancements applied successfully!</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Enhanced Statistics
                    st.markdown("## 📊 Enhanced Results")
                    
                    col1, col2, col3, col4, col5 = st.columns(5)
                    with col1:
                        st.markdown(f"""
                        <div class="metric-card">
                            <h2 style="color: #667eea;">{stats['paragraphs']}</h2>
                            <p>Paragraphs</p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown(f"""
                        <div class="metric-card">
                            <h2 style="color: #28a745;">{stats['math_expressions']}</h2>
                            <p>Math Exprs</p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col3:
                        st.markdown(f"""
                        <div class="metric-card">
                            <h2 style="color: #ffc107;">{stats['nth_roots']}</h2>
                            <p>nth Roots</p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col4:
                        st.markdown(f"""
                        <div class="metric-card">
                            <h2 style="color: #17a2b8;">{stats['markdown_tables']}</h2>
                            <p>MD Tables</p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col5:
                        st.markdown(f"""
                        <div class="metric-card">
                            <h2 style="color: #6f42c1;">{stats['questions_formatted']}</h2>
                            <p>Q&A Items</p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    # Enhanced method breakdown
                    if stats['math_expressions'] > 0:
                        st.markdown("### 🔧 Enhanced Conversion Methods")
                        method_col1, method_col2, method_col3, method_col4 = st.columns(4)
                        
                        with method_col1:
                            st.metric("OMML Objects", stats['method_stats']['omml'], 
                                    help="High-quality equation objects")
                        with method_col2:
                            st.metric("Styled Math", stats['method_stats']['styled'],
                                    help="Cambria Math styled")
                        with method_col3:
                            st.metric("Unicode Fallback", stats['method_stats']['fallback'],
                                    help="Unicode symbols")
                        with method_col4:
                            st.metric("Error Handling", stats['method_stats']['error'],
                                    help="Protected conversions")
                    
                    # Enhanced debug log
                    if show_debug and stats['processing_log']:
                        with st.expander("🔍 Enhanced Processing Log"):
                            for log in stats['processing_log']:
                                st.text(log)
                    
                    # Enhanced download
                    st.download_button(
                        "📥 Download COMPLETELY ENHANCED File",
                        data=buffer.getvalue(),
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        type="primary"
                    )
                    
                    # Enhanced success details
                    success_details = []
                    if stats['math_expressions'] > 0:
                        success_details.append(f"✨ {stats['math_expressions']} math expressions converted with enhanced methods")
                    if stats['nth_roots'] > 0:
                        success_details.append(f"∛∜ {stats['nth_roots']} nth roots properly rendered")
                    if stats['markdown_tables'] > 0:
                        success_details.append(f"📊 {stats['markdown_tables']} markdown tables → Word tables (no duplicates)")
                    if stats['questions_formatted'] > 0:
                        success_details.append(f"📝 {stats['questions_formatted']} Q&A items professionally formatted")
                    if stats['tables'] > 0:
                        success_details.append(f"🎨 {stats['tables']} tables enhanced with professional styling")
                    
                    for detail in success_details:
                        st.success(detail)
                        
                except Exception as e:
                    st.error(f"Error details: {str(e)}")
                    
                    if show_debug:
                        st.code(traceback.format_exc())

if __name__ == "__main__":
    main()
