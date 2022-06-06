from tkinter import *
from tkinter import ttk, messagebox
from contar import *
import sqlite3
import win32print
import datetime


# tabela fechamento
banco = sqlite3.connect(r"c:\FLPrograms\DBSupermercado Hely.sql")
c = banco.cursor()

# tabela gastos(idcontrole, dia, gastosSemRetorno, gastosComRetorno, funcionario)
    
def confirmar():
    
    jnl1 = Tk()
    jnl1.geometry("300x300")
    jnl1.title("Logon")
    
    def verificar():
            users = tnome.get()
            user = users.lower()
            passwd = ekey.get()
            
            with open(r'c:\FLPrograms\users.txt','w') as usuario:
                usuario.write(str(f"{user}"))
            
            c.execute("SELECT supervisor FROM fechamento")
            lista = c.fetchall()
            
            controle = 1
            controle1 = len(lista)
            
            for row in lista:
                
                if user == "felipe " and passwd =="FLProgram":
                    
                    jnl1.destroy()
                    administrador()
                    Programador()
                    break
                
                elif row == (f'{user}',):
                    b = lista.index((f'{user}',))
                   
                    c.execute("SELECT senha FROM fechamento")
                    lista2 = c.fetchall()
                    
                    if lista2[b] == ((f'{passwd}',)):
                        
                        jnl1.destroy()
                        administrador()
                    
                    else:
                        erro = Label(jnl1,text="Acesso Negado", font="arial 13", bg="#ffffff", width=15)
                        erro.pack(padx=15,pady=15)
                        btnfechar = Button(jnl1,text="Fechar",width=5, font="arial 13", height=1, command=jnl1.destroy)
                        btnfechar.pack(padx=20, pady=20)
                        lblobs = Label(jnl1, text='Para informacões de software, \n logar com "INFO". ', font = "arial 10" )
                        lblobs.pack(padx=30, pady=30)
                        break

                    
                
                elif user == "info":
                    
                    jnl2 = Tk()
                    jnl2.geometry("400x300")
                    jnl2.title("INFORMAÇÃO DE SOFTWARE")
                    
                    criador = Label(jnl2,text="FLPrograms", font="Times 12 bold", width=20)
                    criador.pack(anchor=CENTER)
                    versao = Label(jnl2,text="Vesão 1.0/03.2022", font="Times 12 bold", width=20)
                    versao.pack(anchor=CENTER)
                    cntt = Label(jnl2,text="Contato: felipedsrodrigues@outlook.com \n (97)984165908", font="Times 12 bold", width=30)
                    cntt.pack(anchor=CENTER)
                    btnfechar = Button(jnl2,text="Fechar", font="Times 12 bold" , width=10, bg="#ffffff", command=jnl2.destroy)
                    btnfechar.pack(anchor=CENTER)
                
                    jnl2.mainloop                        
                    break
                elif controle == controle1:    
                    
                    erro = Label(jnl1,text="Acesso Negado", font="arial 13", bg="#ffffff", width=15)
                    erro.pack(padx=15,pady=15)
                    btnfechar = Button(jnl1,text="Fechar",width=5, font="arial 13", height=1, command=jnl1.destroy)
                    btnfechar.pack(padx=20, pady=20)
                    lblobs = Label(jnl1, text='Para informacões de software, \n logar com "INFO". ', font = "arial 10" )
                    lblobs.pack(padx=30, pady=30)
                    
                    break

                else:
                    
                    controle += 1
                    
    
    
    nome = Label(jnl1, text="Usuario",font= "arial 13",width=10)
    nome.pack(padx=0 , pady=5,  anchor=CENTER)
    tnome = Entry(jnl1, width=10, background="#ffffff", font="arial 13")
    tnome.pack(padx=5, pady=0, anchor=CENTER)
    Key = Label(jnl1, text="Senha", font="arial 13")
    Key.pack(padx=0, pady=5, anchor=CENTER)
    ekey = Entry(jnl1, width=10,show="*",font="arial 13",background="#ffffff")
    ekey.pack(padx=5, pady=0, anchor=CENTER)
    
    
    btnconfirmar = Button(jnl1,text="Logar",width=5,bd=1,bg="#4dc3ff", command=verificar)
    btnconfirmar.pack(padx=20,pady=20)

    
   
    
    jnl1.mainloop()


def administrador():
    
    jnl2 = Tk()
    jnl2.title("Administrativo")
    jnl2.geometry("400x400")
    jnl2.state("zoomed")
    titulo = Label(jnl2,text="ADIMINISTRATIVO", width=20 ,font="times 17 bold")
    titulo.pack(anchor=CENTER, pady=10)
    
    
    def gastos():
        jnl2.destroy()
        gerirgastos()
    
    fonte = "arial 13"
    btngast = Button(jnl2,text="Gerenciar Gastos", bg="#4dc3ff", width=20 , font=fonte, command=gastos)
    btngast.pack(anchor=E,padx=20,pady=9)
    
    btnrelatorio = Button(jnl2,text="Consultar Relatorios", bg="#4dc3ff",width=20, font=fonte , command=relatorio)
    btnrelatorio.pack(anchor=E,padx=20)
    

                
                
   
            
       
    jnl2.mainloop()

def Programador():
    jnl3 = Tk()
    jnl3.title("Configurações")    
    jnl3.geometry("600x600")
    jnl3.state("zoomed")
    
    todas_impressoras = win32print.EnumPrinters(2)
    lista_impressoras = []
    for row in todas_impressoras:
        impressora = row[2]
        lista_impressoras.append(impressora)
    
    
    def print_padrao():
        jnl4 = Tk()
        lblselect = Label(jnl4,text="Selecione a Impressora padrão", font="Arial 12", anchor=CENTER, width=30)
        lblselect.pack()
        cb_print = ttk.Combobox(jnl4,values=lista_impressoras, font="arial 12",width=20)
        cb_print.pack(anchor=CENTER)
        def select():
            
            with open(r'c:\FLPrograms\Impressora padrao.txt','w') as impressorapadrao:
                impressorapadrao.write(cb_print.get())
                jnl4.destroy()
                
        btnconfirm = Button(jnl4,text="Selecionar e Fechar",width=20,bg="#4dc3ff", command=select )
        btnconfirm.pack(anchor=CENTER)
        
        jnl4.mainloop

    def new_super():
        jnl6 = Tk()
        jnl6.title("Casdastrar")
        jnl6.geometry("500x250")
        
        
        
        
        lblnome = Label(jnl6,text="Nome \n \nSenha", width=10, height=3 , font="arial 12")
        lblnome.place(x=20, y=20)
        
        enome = Entry(jnl6,width=20,font="arial 12")
        enome.place(x=100, y=20)
        
        esenha = Entry(jnl6,width=20, font="arial 12")
        esenha.place(x=100, y=60)
        
        
        def cadastrar():
            nome = enome.get()
            senha = esenha.get()
            
            c.execute("SELECT supervisor FROM fechamento")
            supervisores = c.fetchall()
            control = 1
            control1 = len(supervisores)
            
            for nomes in supervisores:
                if nomes[0] == nome:
                    
                    repetido = messagebox.askyesno(title="Usuario Repetido!", message="Supervisor já cadastrado, \n deseja atualizar nova senha?")
                    if repetido:
                        c.execute(f"UPDATE fechamento SET senha = '{senha}' WHERE supervisor = '{nome}'")
                        banco.commit()
                        jnl6.destroy()
                        break
                    else:
                        pass
                elif control == control1:
                
                    confirma = messagebox.askyesno(title="", message="Confirmar Inclusao de novo Supervisor ?")
                    if confirma:
                        
                        c.execute(f"INSERT INTO fechamento (supervisor , senha) VALUES ('{nome}','{senha}')")
                        banco.commit()
                        jnl6.destroy() 
                        break
                else:
                    control +=1
            
        def apagar():
            nome = enome.get()
            
            c.execute("SELECT supervisor FROM fechamento")
            supervir = c.fetchall()
            control = 1
            control1 = len(supervir)
            for name in supervir:
                if name[0] == nome:
                     
                    askdelete = messagebox.askyesno(title="!", message="APAGAR USUARIO ?")
                    if askdelete:
                        c.execute(f"DELETE FROM fechamento WHERE supervisor = '{nome}' ")
                        banco.commit()
                        jnl6.destroy()
                        break
                elif control == control1 :
                    messagebox.showinfo(title="", message="Usuario não encontrado")
                else:
                    control += 1
               
        btncadastrar = Button(jnl6, text="Cadastrar Novo supervisor", font = "arial 12", width=22,bg="#4dc3ff", command=cadastrar)
        btncadastrar.place(x=50, y= 120)
        
        btndelete = Button(jnl6,text="Apagar Supervisor", font= "arial 12", width=22, bg="red", command=apagar)
        btndelete.place(x=260, y= 120)
        
        jnl6.mainloop()

    btnprint = Button(jnl3,text="Selecionar impressora padrão", font="arial 12 bold",bg="#4dc3ff", width=25, command=print_padrao)
    btnprint.pack(anchor = CENTER, pady=60)
    
    btnsupervisor = Button(jnl3, text="Cadastrar novo Supervisor", font = "arial 12 bold",bg="#4dc3ff", width=25, command=new_super)
    btnsupervisor.pack(anchor=CENTER, pady=2)
    
    btndb = Button(jnl3, text="Cadastrar Servidor Local" , font="arial 12 bold" , bg="#4dc3ff",width=25)
    btndb.pack(anchor=CENTER , pady=2)
    
    jnl3.mainloop()

def gerirgastos():
    
    
   
        
    jnl5 = Tk()
    jnl5.title("Adicionar Gastos")
    jnl5.geometry("1000x800")
    jnl5.state("zoomed")
    
    def voltar():
        jnl5.destroy()
        administrador()
    
    btnvoltar = Button(jnl5,text="Voltar", font="arial 10", command=voltar)
    btnvoltar.place(x=0,y=1)
    
    titulo = Label(jnl5,text="Gerência de gastos", width=20, font="times 18 bold" )
    titulo.place(x=50, y=1)
    
    lblday = Label(jnl5,text="Dia", font="arial 12 bold", bg="#fff" , bd=5 , width=4, anchor=W)
    lblday.place(x=50, y=40)
    edia = Entry(jnl5,width=10, font="arial 12")
    edia.place(x=50,y=70)
    
    datual = datetime.datetime.now()
    mesatual = datual.month
    diaatual = datual.day
    if mesatual < 10 and diaatual < 10 :
        edia.insert(0,'0' + str(datual.day) + "/" + '0' + str(datual.month) + "/" + str(datual.year) )
    elif mesatual < 10:
        edia.insert(0, str(datual.day) + "/" + '0' + str(datual.month) + "/" + str(datual.year) )
    elif diaatual < 10:
        edia.insert(0,'0' + str(datual.day) + "/" + str(datual.month) + "/" + str(datual.year) )
    else:
        edia.insert(0, str(datual.day) + "/" + str(datual.month) + "/" + str(datual.year) )
    
    fonte = "arial 12"
    
    lbldescricao1 = Label(jnl5,text="Descrição gastos sem retorno", width=30, font="arial 12 bold", bg="#fff")
    lbldescricao1.place(x=40, y=100)
    lbldescricao2 = Label(jnl5,text="Descrição gastos com retorno", width=30, font="arial 12 bold", bg="#fff")
    lbldescricao2.place(x=500, y=100)
    
    lbldescricao11 = Label(jnl5,text="valor", width=10, font="arial 12 bold", bg="#fff")
    lbldescricao11.place(x=352, y=100)
    lbldescricao22 = Label(jnl5,text="valor", width=10, font="arial 12 bold", bg="#fff")
    lbldescricao22.place(x=812, y=100)
    
    
    #coluna gastos sem retorno
    
    egastos11 = Entry(jnl5,width=35, font=fonte)
    egastos11.place(x=40, y=120) 
    egastos111 = Entry(jnl5,width=15, font=fonte)
    egastos111.place(x=350, y=120) 
    
    egastos12 = Entry(jnl5,width=35, font=fonte)
    egastos12.place(x=40, y=143) 
    egastos122 = Entry(jnl5,width=15, font=fonte)
    egastos122.place(x=350, y=143) 
    
    egastos13 = Entry(jnl5,width=35, font=fonte)
    egastos13.place(x=40, y=166) 
    egastos133 = Entry(jnl5,width=15, font=fonte)
    egastos133.place(x=350, y=166) 
    
    egastos14 = Entry(jnl5,width=35, font=fonte)
    egastos14.place(x=40, y=189) 
    egastos144 = Entry(jnl5,width=15, font=fonte)
    egastos144.place(x=350, y=189) 
    
    egastos15 = Entry(jnl5,width=35, font=fonte)
    egastos15.place(x=40, y=212) 
    egastos155 = Entry(jnl5,width=15, font=fonte)
    egastos155.place(x=350, y=212) 
    
    egastos16 = Entry(jnl5,width=35, font=fonte)
    egastos16.place(x=40, y=235) 
    egastos166 = Entry(jnl5,width=15, font=fonte)
    egastos166.place(x=350, y=235) 
    
    egastos17 = Entry(jnl5,width=35, font=fonte)
    egastos17.place(x=40, y=258) 
    egastos177 = Entry(jnl5,width=15, font=fonte)
    egastos177.place(x=350, y=258) 
    
    egastos18 = Entry(jnl5,width=35, font=fonte)
    egastos18.place(x=40, y=281) 
    egastos188 = Entry(jnl5,width=15, font=fonte)
    egastos188.place(x=350, y=281) 
    
    egastos19 = Entry(jnl5,width=35, font=fonte)
    egastos19.place(x=40, y=304) 
    egastos199 = Entry(jnl5,width=15, font=fonte)
    egastos199.place(x=350, y=304) 
    
    egastos10 = Entry(jnl5,width=35, font=fonte)
    egastos10.place(x=40, y=327) 
    egastos100 = Entry(jnl5,width=15, font=fonte)
    egastos100.place(x=350, y=327) 
    
    # coluna gastos com retorno
    
    egastos21 = Entry(jnl5,width=35, font=fonte)
    egastos21.place(x=500, y=120) 
    egastos211 = Entry(jnl5,width=15, font=fonte)
    egastos211.place(x=810, y=120) 
    
    egastos22 = Entry(jnl5,width=35, font=fonte)
    egastos22.place(x=500, y=143) 
    egastos222 = Entry(jnl5,width=15, font=fonte)
    egastos222.place(x=810, y=143) 

    egastos23 = Entry(jnl5,width=35, font=fonte)
    egastos23.place(x=500, y=166) 
    egastos233 = Entry(jnl5,width=15, font=fonte)
    egastos233.place(x=810, y=166) 
    
    egastos24 = Entry(jnl5,width=35, font=fonte)
    egastos24.place(x=500, y=189) 
    egastos244 = Entry(jnl5,width=15, font=fonte)
    egastos244.place(x=810, y=189) 
    
    egastos25 = Entry(jnl5,width=35, font=fonte)
    egastos25.place(x=500, y=212) 
    egastos255 = Entry(jnl5,width=15, font=fonte)
    egastos255.place(x=810, y=212) 
    
    egastos26 = Entry(jnl5,width=35, font=fonte)
    egastos26.place(x=500, y=235) 
    egastos266 = Entry(jnl5,width=15, font=fonte)
    egastos266.place(x=810, y=235) 
    
    egastos27 = Entry(jnl5,width=35, font=fonte)
    egastos27.place(x=500, y=281) 
    egastos277 = Entry(jnl5,width=15, font=fonte)
    egastos277.place(x=810, y=281) 
    
    egastos28 = Entry(jnl5,width=35, font=fonte)
    egastos28.place(x=500, y=304) 
    egastos288 = Entry(jnl5,width=15, font=fonte)
    egastos288.place(x=810, y=304) 
    
    egastos29 = Entry(jnl5,width=35, font=fonte)
    egastos29.place(x=500, y=327) 
    egastos299 = Entry(jnl5,width=15, font=fonte)
    egastos299.place(x=810, y=327) 
    
    egastos20 = Entry(jnl5,width=35, font=fonte)
    egastos20.place(x=500, y=258) 
    egastos200 = Entry(jnl5,width=15, font=fonte)
    egastos200.place(x=810, y=258) 
    
    
    
    lbltotalsretorno = Label(jnl5, text="Total:",font= fonte, width=15, bg="white")
    lbltotalsretorno.place(x=230 , y=350)
    lbltotalsretornov = Label(jnl5, text="",font= fonte, width=15, bg="white")
    lbltotalsretornov.place(x=350, y=350)
    lbltotalcretorno = Label(jnl5, text="Total:",font= fonte, width=15, bg="white")
    lbltotalcretorno.place(x=670 , y=350)
    lbltotalcretornov = Label(jnl5, text="",font= fonte, width=15, bg="white")
    lbltotalcretornov.place(x=810, y=350)
    
    def calcular():
        w = [egastos11.get(),egastos12.get(),egastos13.get(),egastos14.get(),egastos15.get(),egastos16.get(),egastos17.get(),egastos18.get(),egastos19.get(),egastos10.get()]
        x = [egastos111.get(),egastos122.get(),egastos133.get(),egastos144.get(),egastos155.get(),egastos166.get(),egastos177.get(),egastos188.get(),egastos199.get(),egastos100.get()]
        y = [egastos21.get(),egastos22.get(),egastos23.get(),egastos24.get(),egastos25.get(),egastos26.get(),egastos27.get(),egastos28.get(),egastos29.get(),egastos20.get()]
        z = [egastos211.get(),egastos222.get(),egastos233.get(),egastos244.get(),egastos255.get(),egastos266.get(),egastos277.get(),egastos288.get(),egastos299.get(),egastos200.get()]
        
        descricaogastossr = []
        valoresgastossr = []
        descricaogastoscr = []
        valoresgastoscr = []
        
        ttsrt = 0
        ttcrt = 0
        for num in range(0,10):
            
            
            
            
            if w[num] != "" and x[num] != "":
                ttsrt += float(x[num].replace(",","."))
                descricaogastossr.append(w[num])
                valoresgastossr.append(x[num])
                
            else:
                pass
            if y[num] != "" and z[num] != "":
                ttcrt += float(z[num].replace(",","."))
                descricaogastoscr.append(y[num])
                valoresgastoscr.append(z[num])
                
            else:
                pass
            
            
            if w[num] != "" and x[num] == "":
                messagebox.showinfo(title="ERRO", message=f'ITEM "{w[num]}" SEM VALOR DESCRITO')
            else:
                
                pass
            if y[num] != "" and z[num] == "":
                messagebox.showinfo(title="ERRO", message=f'ITEM "{y[num]}" SEM VALOR DESCRITO')
            else:
                
                pass
            
            if w[num] == "" and x[num] != "":
                messagebox.showinfo(title="ERRO", message=f'ITEM "{x[num]}" SEM DESCRIÇÃO')
            else:
                pass
            if y[num] == "" and z[num] != "":
                messagebox.showinfo(title="ERRO", message=f'ITEM "{z[num]}" SEM VALOR DESCRIÇÃO')
            else:    
                pass
            
            
        lbltotalsretornov["text"] = ttsrt
        lbltotalcretornov["text"] = ttcrt
        
        
        def enviar():
            
            #c.close()
            banco2 = sqlite3.connect(r"c:\FLPrograms\DBgastosSupermercado Hely.sql")
            cursor = banco2.cursor()
        
            confirmacao = messagebox.askyesno(title="Confirmar", message="Confirmar Envio ?")
            if confirmacao:
                
                with open (r"c:\FLPrograms\users.txt","r") as users:
                    for usuarios in users:
                        usuario = usuarios
                
                
                diaa = edia.get()
                
                ts = len(descricaogastossr)
                if ts > 0:
                    for rows in range(0,ts):
                        dgs = descricaogastossr[rows]
                        vsr = valoresgastossr[rows]
                        cursor.execute(f"INSERT INTO gastos (dia , gastosSemRetorno , valorSR,funcionario) VALUES ('{diaa}', '{dgs}', {vsr} , '{usuario}')")
                        banco2.commit()
                    
                tc = len(descricaogastoscr)
                if tc > 0 :
                    for value in range(0,tc):
                        dgc = descricaogastoscr[rows]
                        vcr = valoresgastoscr[rows]
                        cursor.execute(f"INSERT INTO gastos (dia,gastosComRetorno,valorCR, funcionario) VALUES ('{diaa}', '{dgc}' , {vcr} , '{usuario}')")
                        banco2.commit()
                        
                jnl5.destroy()
                gerirgastos()  
        
        btnenviar = Button(jnl5,text="Enviar", width=10, bg="#eee", font=fonte, command=enviar)
        btnenviar.place(x=750, y=390)
    
    btnconfirm = Button(jnl5,text="Confirmar", width=10, bg="#4dc3ff", font=fonte, command=calcular)
    btnconfirm.place(x= 850, y = 390)
    
    lista = ["egastos1, "]    
        
    
    
    
    jnl5.mainloop()


def relatorsdcio():
    wb = load_workbook(r"c:\FLPrograms\r22m233.xlsx")
    
    sh = wb['Table']
    
    sh['b6'].value = 200
    
    print(sh['b6'].value)
    wb.save(filename="Março2022.xlsx")

#Programador()