import tkinter as tk
from tkinter import filedialog, messagebox, font, ttk
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH

class HTMLConverter:
    def __init__(self, master):
        self.master = master
        self.master.title("Conversor Avanzado a HTML")
        self.setup_ui()
        self.setup_tags()

    def setup_ui(self):
        # Widgets principales
        ttk.Button(self.master, text="Convertir DOCX", command=self.convert_docx).pack(pady=5)
        
        self.input_text = tk.Text(self.master, wrap=tk.WORD, height=15, undo=True)
        self.input_text.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        
        self.setup_toolbar()
        
        ttk.Button(self.master, text="Convertir Texto", command=self.convert_text).pack(pady=5)
        
        self.output_text = tk.Text(self.master, wrap=tk.WORD, height=15)
        self.output_text.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

    def setup_toolbar(self):
        # Barra de herramientas con estilos
        toolbar = ttk.Frame(self.master)
        toolbar.pack(fill=tk.X)
        
        styles = [
            ('bold', 'Negrita', self.toggle_style),
            ('italic', 'Cursiva', self.toggle_style),
            ('underline', 'Subrayado', self.toggle_style)
        ]
        
        for style, text, cmd in styles:
            btn = ttk.Button(toolbar, text=text, command=lambda s=style: cmd(s))
            btn.pack(side=tk.LEFT, padx=2)
        
        # Selector de alineación
        self.align_var = tk.StringVar(value='left')
        align_menu = ttk.OptionMenu(toolbar, self.align_var, 'left', 'left', 'center', 'right', 'justify',
                                   command=self.apply_alignment)
        align_menu.pack(side=tk.RIGHT, padx=5)

    def setup_tags(self):
        # Configuración de tags para estilos
        base_font = font.Font(font=self.input_text['font'])
        self.input_text.tag_configure('bold', font=(base_font.actual()['family'], base_font.actual()['size'], 'bold'))
        self.input_text.tag_configure('italic', font=(base_font.actual()['family'], base_font.actual()['size'], 'italic'))
        self.input_text.tag_configure('underline', underline=True)
        
        # Tags para alineación
        for align in ['left', 'center', 'right', 'justify']:
            self.input_text.tag_configure(align, justify=align)

    def toggle_style(self, style):
        # Alternar estilos en texto seleccionado
        current_tags = self.input_text.tag_names(tk.SEL_FIRST)
        if style in current_tags:
            self.input_text.tag_remove(style, tk.SEL_FIRST, tk.SEL_LAST)
        else:
            self.input_text.tag_add(style, tk.SEL_FIRST, tk.SEL_LAST)

    def apply_alignment(self, alignment):
        # Aplicar alineación al párrafo actual
        line_start = self.input_text.index(tk.INSERT + " linestart")
        line_end = self.input_text.index(tk.INSERT + " lineend")
        self.input_text.tag_add(alignment, line_start, line_end)

    def get_html_styles(self, index):
        # Obtener estilos aplicados en una posición
        tags = self.input_text.tag_names(index)
        styles = []
        
        if 'bold' in tags: styles.append('font-weight:bold')
        if 'italic' in tags: styles.append('font-style:italic')
        if 'underline' in tags: styles.append('text-decoration:underline')
        
        return '; '.join(styles)

    def convert_text(self):
        # Convertir texto con estilos a HTML
        html = []
        text_content = self.input_text.get(1.0, tk.END).split('\n')
        
        for line in text_content[:-1]:  # Ignorar última línea vacía
            if not line.strip():
                html.append('<p>&nbsp;</p>')
                continue
                
            line_html = []
            pos = self.input_text.search(line, 1.0, stopindex=tk.END)
            
            while pos:
                end_pos = f"{pos}+{len(line)}c"
                styles = self.get_html_styles(pos)
                line_html.append(f'<span style="{styles}">{line}</span>')
                pos = self.input_text.search(line, end_pos, stopindex=tk.END)
                
            alignment = self.input_text.tag_names(f"{self.input_text.index(line + '.0')} linestart")[0]
            html.append(f'<p style="text-align:{alignment}">{"".join(line_html)}</p>')
        
        self.output_text.delete(1.0, tk.END)
        self.output_text.insert(tk.END, '\n'.join(html))

    def convert_docx(self):
        # Convertir documento Word a HTML
        file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if not file_path: return
        
        try:
            doc = docx.Document(file_path)
            html = []
            
            for para in doc.paragraphs:
                para_html = []
                alignment = self.get_docx_alignment(para.alignment)
                
                for run in para.runs:
                    text = run.text.strip()
                    if not text: continue
                    
                    styles = []
                    if run.bold: styles.append('font-weight:bold')
                    if run.italic: styles.append('font-style:italic')
                    if run.underline: styles.append('text-decoration:underline')
                    
                    style_attr = f' style="{"; ".join(styles)}"' if styles else ''
                    para_html.append(f'<span{style_attr}>{text}</span>')
                
                if para_html:
                    html.append(f'<p style="text-align:{alignment}">{" ".join(para_html)}</p>')
                else:
                    html.append('<p>&nbsp;</p>')
            
            output_file = file_path.replace('.docx', '_converted.html')
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write('\n'.join(html))
            
            messagebox.showinfo("Éxito", f"Archivo guardado en:\n{output_file}")
        
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo convertir el archivo:\n{str(e)}")

    def get_docx_alignment(self, alignment):
        # Mapear alineación de Word a CSS
        return {
            WD_ALIGN_PARAGRAPH.LEFT: 'left',
            WD_ALIGN_PARAGRAPH.CENTER: 'center',
            WD_ALIGN_PARAGRAPH.RIGHT: 'right',
            WD_ALIGN_PARAGRAPH.JUSTIFY: 'justify'
        }.get(alignment, 'left')

if __name__ == "__main__":
    root = tk.Tk()
    app = HTMLConverter(root)
    root.mainloop()
