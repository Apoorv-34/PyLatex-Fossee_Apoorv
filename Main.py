import pandas as pd
import numpy as np
from pylatex import Document, Section, Subsection, Figure, Tabular, Command, Package
from pylatex.utils import NoEscape, italic
import os

def generate_beam_report():
    print("📊 Reading beam data from Excel...")
    df = pd.read_excel('beam_data.xlsx')
    
    # article class with 11pt for a dense, professional technical look
    doc = Document(documentclass='article', document_options=['a4paper', '11pt'])
    
    # --- Specialized Formatting Packages ---
    doc.packages.append(Package('tgadventor')) # Clean, geometric sans-serif font
    doc.packages.append(Package('tcolorbox'))  # For technical summary boxes
    doc.packages.append(Package('xcolor'))
    doc.packages.append(Package('tikz'))
    doc.packages.append(Package('pgfplots'))
    doc.packages.append(Package('geometry', options=['margin=0.75in']))
    doc.packages.append(Package('fancyhdr'))

    # --- Custom Theme & Preamble ---
    doc.preamble.append(NoEscape(r'\definecolor{slate}{HTML}{273746}'))
    doc.preamble.append(NoEscape(r'\definecolor{accent}{HTML}{2E86C1}'))
    doc.preamble.append(NoEscape(r'\renewcommand{\familydefault}{\sfdefault}'))
    doc.preamble.append(NoEscape(r'\pgfplotsset{compat=1.18}'))

    # --- Page 1: Technical Cover & Overview ---
    doc.append(NoEscape(r'\begin{center}'))
    doc.append(NoEscape(r'{\Huge \textbf{\color{slate}TECHNICAL MEMORANDUM} \par}'))
    doc.append(NoEscape(r'\vspace{0.5cm}'))
    doc.append(NoEscape(r'{\large \textit{Structural Analysis of Flexural Members}} \par'))
    doc.append(NoEscape(r'\vspace{0.8cm}'))
    
    # Summary Info Box
    doc.append(NoEscape(r'''
    \begin{tcolorbox}[colframe=slate, colback=white, title=Project Identifiers]
    \textbf{Lead Engineer:} Apoorv Goyal \\
    \textbf{Registration ID:} 23BCG10116 \\
    \textbf{Subject:} Simply Supported Beam - 12m Span Analysis \\
    \textbf{Date of Issue:} \today
    \end{tcolorbox}
    '''))
    doc.append(NoEscape(r'\end{center}'))
    doc.append(NoEscape(r'\vspace{0.5cm}'))

    with doc.create(Section('Project Scope')):
        doc.append('This memorandum presents the calculated internal force distribution for a primary structural element. The analysis focuses on deriving the Shear Force Diagram (SFD) and Bending Moment Diagram (BMD) under static load conditions.')
        
        with doc.create(Subsection('Configuration Modeling')):
            doc.append('The beam is modeled with ideal pinned-roller boundary conditions. The geometry and loading path are visualized in the figure below:')
            with doc.create(Figure(position='h!')) as beam_fig:
                if os.path.exists('beam.png'):
                    beam_fig.add_image('beam.png', width=NoEscape(r'0.6\textwidth'))
                beam_fig.add_caption('Analytical Free Body Diagram.')

    # Force a page break to make it 2 pages
    doc.append(NoEscape(r'\newpage'))

    # --- Page 2 Setup ---
    doc.preamble.append(NoEscape(r'\pagestyle{fancy}'))
    doc.preamble.append(NoEscape(r'\fancyhf{}'))
    doc.preamble.append(NoEscape(r'\lfoot{\small \color{gray} Apoorv Goyal | 23BCG10116}'))
    doc.preamble.append(NoEscape(r'\rfoot{\small Page \thepage}'))

    with doc.create(Section('Numerical Computation Matrix')):
        doc.append('Sampled data points extracted from the finite element simulation:')
        doc.append(NoEscape(r'\vspace{0.3cm}'))
        
        # Clean, modern table with colored header row
        with doc.create(Tabular('p{3cm} p{3cm} p{3cm}', row_height=1.4)) as table:
            table.add_row((NoEscape(r'\textbf{Point (m)}'), NoEscape(r'\textbf{Shear (kN)}'), NoEscape(r'\textbf{Moment (kNm)}')))
            table.add_hline()
            for _, row in df.iloc[::2].iterrows():
                table.add_row((f"{row['X']:.2f}", f"{row['Shear force']:.1f}", f"{row['Bending Moment']:.1f}"))

    with doc.create(Section('Internal Force Envelopes')):
        doc.append('The graphical plots below describe the mechanical response of the beam.')
        
        with doc.create(Subsection('Shear Force Variance')):
            doc.append(NoEscape(generate_sfd_plot(df)))

        with doc.create(Subsection('Bending Moment Variance')):
            doc.append(NoEscape(generate_bmd_plot(df)))

    # Generate PDF
    print("📄 Generating Memorandum PDF...")
    doc.generate_pdf('Apoorv_Goyal_Memorandum', clean_tex=False, compiler='pdflatex')
    print("✅ Success: Apoorv_Goyal_Memorandum.pdf")

def generate_sfd_plot(df):
    coords = "".join([f"({r['X']}, {r['Shear force']}) " for _, r in df.iterrows()])
    return r"""
    \begin{center}
    \begin{tikzpicture}
    \begin{axis}[width=11cm, height=5cm, axis x line=center, axis y line=left, 
        xlabel=$x$, ylabel=$V$, grid=major, grid style={dotted, slate!30},
        title=\textbf{V-Diagram}, title style={color=accent}]
    \addplot[thick, accent, fill=accent, fill opacity=0.1] coordinates {""" + coords + r"""} \closedcycle;
    \end{axis}
    \end{tikzpicture}
    \end{center}
    """

def generate_bmd_plot(df):
    coords = "".join([f"({r['X']}, {r['Bending Moment']}) " for _, r in df.iterrows()])
    return r"""
    \begin{center}
    \begin{tikzpicture}
    \begin{axis}[width=11cm, height=5cm, axis x line=bottom, axis y line=left, 
        xlabel=$x$, ylabel=$M$, grid=major, grid style={dotted, slate!30},
        title=\textbf{M-Diagram}, title style={color=slate}]
    \addplot[thick, slate, fill=slate, fill opacity=0.15] coordinates {""" + coords + r"""} \closedcycle;
    \end{axis}
    \end{tikzpicture}
    \end{center}
    """

if __name__ == "__main__":
    generate_beam_report()