import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from glob import glob
import re
import tempfile
import sys

try: 
    from PyPDF2 import PdfFileWriter, PdfFileReader
except:
    messagebox.showerror(":-(","The PyPDF2 isn't installed, but we will try install it now...")
    try:
        subprocess.call(f'{sys.executable} -m pip install PyPDF2')
        from PyPDF2 import PdfFileWriter, PdfFileReader
    except:
        messagebox.showerror(':-(', "It didn't work. You can download PyPDF2-1.26.0-py2.py3-none-any.whl, put in the same path os this file and run again")
        import sys
        sys.exit()

def executar():
    # Get form data
    folder = eDiretorio.get()
    senha = eSenha.get()
    rot = int(cbRotacao.get())
    merge = ckMerge.state()
    fl0 = ePaginas.get()
    try:
        chunklimit = float(cbLimite.get().replace(',','.'))
    except:
        chunklimit = False

    fl = []  # Array de folhas do documento final

    if folder:
        # Caso haja um diretírio informado ler todos os pdf deste diretório
        docs = glob('{}/*.pdf'.format(folder))
        _nsre = re.compile('([0-9]+)')
        def natural_sort_key(s): # função para considerar números não como string em ordenamento
            return [int(text) if text.isdigit() else text.lower()
                    for text in re.split(_nsre, s)]
        docs.sort(key=natural_sort_key)
    else:
        docs = [eArquivo.get()]

    # Criar array de folhas do documento de saída
    if not fl0:
        fl1= 'todas'
    else:
        fl1 = fl0.split(',')
        for tr in fl1:
            if '-' in str(tr):
                fl += list(range(int(tr.split('-')[0]), int(tr.split('-')[1])+1))
            else:
                fl.append(tr)
    gerado = []
    if merge:
        pdfsaida = PdfFileWriter()
        saida = '{}/{}_res.pdf'.format(folder, folder.split('/')[-1].split('\\')[-1])

    for doc in docs:
        if not merge: pdfsaida = PdfFileWriter()
        try:
            pdffonte = PdfFileReader(open(doc, 'rb'))
        except:
            try:
                pdffonte = PdfFileReader(open(doc, 'rb'))
            except:
                messagebox.showerror('Error', f'Not possible to load the file {doc}.')
                return

        try:
            if pdffonte.isEncrypted:
                decript = pdffonte.decrypt(senha)
                if not decript:
                    messagebox.showerror('File using password',
                                 'The pdf file need a password.')
                    return

            # Gerar documento de saída com folhas definidas na variável "fl"
            if fl1 == 'todas':
                fl = range(1, pdffonte.getNumPages() + 1)
            for f in fl:
                pdfsaida.addPage(pdffonte.getPage(int(f)-1).rotateCounterClockwise(rot))

        except:
            messagebox.showerror('Erro ao tentar separar páginas definidas',
                                 'Verifique se usou a forma correta de definir as páginas ("," para separar as páginas/intervalo e "-" para definir intervalo)')
            return

        if not merge:
            saida = '{}_res.pdf'.format(doc[:-4])

        try:
            chunklimit *= 1024*1024
            if chunklimit and not merge:
                tmppages = []
                for n in range(len(fl)):
                    output = PdfFileWriter()
                    output.addPage(pdffonte.getPage(n))
                    fd, temp_file_name = tempfile.mkstemp()
                    outputStream = open(temp_file_name, 'wb')
                    output.write(outputStream)
                    pagesize = os.path.getsize(temp_file_name)
                    outputStream.close()
                    tmppages.append(pagesize)

                siz = 0
                filenumb = 1
                pdftmp1 = PdfFileWriter()
                for p, s in enumerate(tmppages):
                    siz += s
                    if siz < chunklimit:
                        pdftmp1.addPage(pdfsaida.getPage(p))
                    if siz > chunklimit:
                        saida1  = saida[:-4] + '_' + str(filenumb) + '.pdf'
                        with open(saida1, 'wb') as ArqSaida:
                            pdftmp1.write(ArqSaida)
                            gerado.append(saida1)
                        del pdftmp1
                        pdftmp1 = PdfFileWriter()
                        pdftmp1.addPage(pdfsaida.getPage(p))
                        siz = s
                        filenumb += 1
                    if p == len(tmppages) - 1:
                        saida1  = saida[:-4] + '_' + str(filenumb) + '.pdf'
                        with open(saida1, 'wb') as ArqSaida:
                            pdftmp1.write(ArqSaida)
                            gerado.append(saida1)

            elif not merge:
                with open(saida, 'wb') as ArqSaida:
                    pdfsaida.write(ArqSaida)
                    gerado.append(saida)
        except:
            messagebox.showerror('Erro', 'Não foi passível salvar arquivo.')
            return

    try:
        if merge:
            with open(saida, 'wb') as ArqSaida:
                pdfsaida.write(ArqSaida)
                gerado.append(saida)
    except:
        messagebox.showerror('Erro', 'Não foi passível salvar arquivo.')
        return

    plural = 's' if len(gerado)>1 else ''
    messagebox.showinfo('', 'Arquivo{1} gerado{1}:\n{0}'.format('\n'.join(gerado), plural))


def selarq():
    temp = filedialog.askopenfilename(
        filetypes=[('Arquivo PDF', 'pdf')],
        initialdir=os.getcwd())
    if temp:
        eArquivo.delete(0, tk.END)
        eArquivo.insert(0, temp)


def seldir():
    temp = filedialog.askdirectory()
    if temp:
        eDiretorio.delete(0, tk.END)
        eDiretorio.insert(0, temp)


########## Form ##########
root = tk.Tk()
root.resizable(width=tk.FALSE, height=tk.FALSE)
root.title('PDF Select')

ttk.Label(root, text='File:').grid(row=0, column=0, padx=0, pady=3)
eArquivo = ttk.Entry(root, width=50)
eArquivo.grid(row=0, column=1, columnspan=2)
ttk.Button(root, text='Select', command=selarq).grid(row=0, column=3)

ttk.Label(root, text='Path:').grid(row=1, column=0, padx=0, pady=3)
eDiretorio = ttk.Entry(root, width=50)
eDiretorio.grid(row=1, column=1, columnspan=2)
ttk.Button(root, text='Select', command=seldir).grid(row=1, column=3)

ttk.Label(root, text='Pages:').grid(row=2, column=0, padx=0, pady=3)
ePaginas = ttk.Entry(root, width=25)
ePaginas.grid(row=2, column=1, sticky=tk.W)
ttk.Label(root, text='"," for list pagas and "-" for intervals.\nIf not set, all pages are considered.',
          font=('Helvetica', 7)).grid(row=2, column=2, columnspan=2, padx=0, pady=0, sticky=tk.W)

ttk.Label(root, text='Password:').grid(row=3, column=0, padx=3, pady=3)
eSenha = ttk.Entry(root, width=25)
eSenha.grid(row=3, column=1, sticky=tk.W)
ttk.Label(root, text='PDF password.',
          font=('Helvetica', 7)).grid(row=3, column=2, columnspan=2, padx=0, pady=0, sticky=tk.W)

ttk.Label(root, text='Rotation:').grid(row=4, column=0, padx=0, pady=3)
cbRotacao = ttk.Combobox(root, width=5, justify='center')
cbRotacao['values'] = ('0', '90', '-90', '180')
cbRotacao.current(0)
cbRotacao.grid(row=4, column=1, sticky=tk.W)
ttk.Label(root, text='Anticlockwise.',
          font=('Helvetica', 7)).grid(row=4, column=2, columnspan=2, padx=0, pady=3, sticky=tk.W)

ttk.Label(root, text='Limit size:').grid(row=5, column=0, padx=0, pady=3)
cbLimite = ttk.Entry(root, width=8, justify='center')
cbLimite.grid(row=5, column=1, sticky=tk.W)
ttk.Label(root, text='Size limit in MB.',
          font=('Helvetica', 7)).grid(row=5, column=2, columnspan=2, padx=0, pady=3, sticky=tk.W)

merge = tk.IntVar()
ckMerge = ttk.Checkbutton(root, width=25, text='Marge pages', variable=merge)
ckMerge.grid(row=6, column=1, sticky=tk.W)
ttk.Label(root, text='Join all pdf files in one file.',
          font=('Helvetica', 7)).grid(row=6, column=2, columnspan=2, padx=0, pady=10, sticky=tk.W)

ttk.Button(root, text='\t\t\t Run \t\t\t', command=executar).grid(row=7, column=0, columnspan=4, padx=3, pady=6)

root.mainloop()
