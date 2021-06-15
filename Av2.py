import tkinter as tk
from tkinter import ttk
import pandas as pd
from tkinter import messagebox

class PrincipalRAD:
    def __init__(self, win):
        #componentes
        self.lblNome=tk.Label(win, text='Nome do Aluno:')
        self.lblMateria=tk.Label(win, text='Matéria:')
        self.lblNota1=tk.Label(win, text='Nota da Av1:')
        self.lblNota2=tk.Label(win, text='Nota da Av2:')
        self.lblNota3=tk.Label(win, text='Nota da Av3:')
        self.lblNota4=tk.Label(win, text='Nota da Avd:')
        self.lblMedia=tk.Label(win, text='Média')
        self.txtNome=tk.Entry(bd=2)
        self.txtMateria=tk.Entry()
        self.txtNota1=tk.Entry()
        self.txtNota2=tk.Entry()    
        self.txtNota3=tk.Entry()
        self.txtNota4=tk.Entry()    
        self.btnCalcular=tk.Button(win, text='Cadastrar', command=self.fCalcularMedia)  

        #----- Componente TreeView --------------------------------------------

        self.dadosColunas = ("Aluno","Matéria" ,"Nota1", "Nota2","Nota3","Nota4" ,"Média", "Situação")            
        
        
        self.treeMedias = ttk.Treeview(win, 
                                       columns=self.dadosColunas,
                                       selectmode='browse')
        
        self.verscrlbar = ttk.Scrollbar(win,
                                        orient="vertical",
                                        command=self.treeMedias.yview)
        
        self.verscrlbar.pack(side ='right', fill ='x')

        
        self.treeMedias.configure(yscrollcommand=self.verscrlbar.set)
        
        self.treeMedias.heading("Aluno", text="Aluno")
        self.treeMedias.heading("Matéria", text="Matéria")
        self.treeMedias.heading("Nota1", text="Av1")
        self.treeMedias.heading("Nota2", text="Av2")
        self.treeMedias.heading("Nota3", text="Av3")
        self.treeMedias.heading("Nota4", text="Avd")
        self.treeMedias.heading("Média", text="Média")
        self.treeMedias.heading("Situação", text="Situação")
        

        self.treeMedias.column("Aluno",minwidth=0,width=70)
        self.treeMedias.column("Matéria",minwidth=0,width=70)
        self.treeMedias.column("Nota1",minwidth=0,width=70)
        self.treeMedias.column("Nota2",minwidth=0,width=70)
        self.treeMedias.column("Nota3",minwidth=0,width=70)
        self.treeMedias.column("Nota4",minwidth=0,width=70)
        self.treeMedias.column("Média",minwidth=0,width=70)
        self.treeMedias.column("Situação",minwidth=0,width=70)

        self.treeMedias.pack(padx=10, pady=10)
                
        #---------------------------------------------------------------------        
        #posicionamento dos componentes na janela
        #--------------------------------------------------------------------- 
               
        self.lblNome.place(x=100, y=50)
        self.txtNome.place(x=200, y=50)

        self.lblMateria.place(x=100, y=80)
        self.txtMateria.place(x=200, y=80)
        
        self.lblNota1.place(x=100, y=110)
        self.txtNota1.place(x=200, y=110)
        
        self.lblNota2.place(x=100, y=140)
        self.txtNota2.place(x=200, y=140)

        self.lblNota3.place(x=100, y=170)
        self.txtNota3.place(x=200, y=170)

        self.lblNota4.place(x=100, y=200)
        self.txtNota4.place(x=200, y=200)
               
        self.btnCalcular.place(x=100, y=250)
           
        self.treeMedias.place(x=100, y=300)
        self.verscrlbar.place(x=850, y=300, height=225)
        
        
        
        self.id = 0
        self.iid = 0
        
        self.carregarDadosIniciais()

#-----------------------------------------------------------------------------

    def carregarDadosIniciais(self):
        try:
          fsave = 'AlunosCadastrados.xlsx'
          dados = pd.read_excel(fsave)
          print("************ dados dsponíveis ***********")        
          print(dados)
        
          u=dados.count()
          print('u:'+str(u))
          nn=len(dados['Aluno'])          
          for i in range(nn):                        
            nome = dados['Aluno'][i]
            materia = dados['Matéria'][i]
            nota1 = str(dados['Av1'][i])
            nota2 = str(dados['Av2'][i])
            nota3 = str(dados['Av3'][i])
            nota4 = str(dados['Avd'][i])
            media=str(dados['Média'][i])
            situacao=dados['Situação'][i]
                        
            self.treeMedias.insert('', 'end',
                                   iid=self.iid,                                   
                                   values=(nome,materia,
                                           nota1,
                                           nota2,nota3,nota4,
                                           media,
                                           situacao))
            
            
            self.iid = self.iid + 1
            self.id = self.id + 1
        except:
          print('Ainda não existem dados para carregar')
            
#-----------------------------------------------------------------------------
#Salvar dados para uma planilha excel
#-----------------------------------------------------------------------------   
        
    def fSalvarDados(self):
      try:          
        fsave = 'AlunosCadastrados.xlsx'
        dados=[]
        
        
        for line in self.treeMedias.get_children():
          lstDados=[]
          for value in self.treeMedias.item(line)['values']:
              lstDados.append(value)
              
          dados.append(lstDados)
          
        df = pd.DataFrame(data=dados,columns=self.dadosColunas)
        
        planilha = pd.ExcelWriter(fsave)
        df.to_excel(planilha, 'Inconsistencias', index=False)                
        
        planilha.save()
        print('Dados salvos')
      except:
       print('Não foi possível salvar os dados')   
        
        
#-----------------------------------------------------------------------------
#calcula a média e verifica qual é a situação do aluno
#-----------------------------------------------------------------------------
          
    def fVerificarSituacao(self, nota1, nota2, nota3, nota4):
          media=(nota1+nota2+nota4)/3
          if(media>=6.0):
             situacao = 'Aprovado'
          elif(media<6.0):
              situacao = 'Em Recuperação'

          if(nota3 + nota1 >= 6.0 or nota3 + nota2>= 6.0):
                  situacao = 'Aprovado'

          else:
                situacao = 'Reprovado'
            

          return media, situacao
        
#-----------------------------------------------------------------------------
#Imprime os dados do aluno
#-----------------------------------------------------------------------------
          
    def fCalcularMedia(self):
        try:
          nome = self.txtNome.get()
          materia = self.txtMateria.get()
          nota1=float(self.txtNota1.get())
          nota2=float(self.txtNota2.get())
          nota3=float(self.txtNota3.get())
          nota4=float(self.txtNota4.get())
          media, situacao = self.fVerificarSituacao(nota1, nota2, nota3, nota4)
                    
          
          self.treeMedias.insert('', 'end', 
                                 iid=self.iid,                                  
                                 values=(nome,materia, 
                                         str(nota1),
                                         str(nota2),str(nota3),str(nota4),
                                         str(media),
                                         situacao))
          
          
          self.iid = self.iid + 1
          self.id = self.id + 1
          
          self.fSalvarDados()
        except ValueError:
            messagebox.showinfo("Information", "Entre com valores válidos.")
            #eu nao consegui validar, dei uma pesquisada, mais as coisas que eu achava nao encaixa aqui
            # nao sei se foi porque eu nao tenho noçao ou porque realmente nao se encaixa aqui   
        finally:
          self.txtNome.delete(0, 'end')
          self.txtMateria.delete(0, 'end')
          self.txtNota1.delete(0, 'end')
          self.txtNota2.delete(0, 'end')
          self.txtNota3.delete(0, 'end')
          self.txtNota4.delete(0, 'end')

#-----------------------------------------------------------------------------
#Programa Principal
#-----------------------------------------------------------------------------          

janela=tk.Tk()
principal=PrincipalRAD(janela)
janela.title('Bem Vindo ao Cadastro de Matéria e Notas.')
janela.geometry("890x650+10+10")
janela.mainloop()

#-----------------------------------------------------------------------------
#Alvaro Pereira Da Silva
#202001188827
#-----------------------------------------------------------------------------
