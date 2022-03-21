'''
 Cambios de Inventario
'''
from tkinter import Tk, Toplevel, messagebox
from datetime import date
from A5 import A5, A5Tk, A5xlsX, UpdateChk
import win32com.client
import win32timezone
import win32ui
import tempfile
import os
import re
import traceback
def show_error(self, *args):
    err = traceback.format_exception(*args)
    messagebox.showerror('Exception',err)
Tk.report_callback_exception = show_error

#import ctypes
#ctypes.windll.user32.ShowWindow( ctypes.windll.kernel32.GetConsoleWindow(), 6 )
EcoLogo= '''
    R0lGODlhEAAQAHAAACwAAAAAEAAQAIfGxsbFxcXExMTDw8PCwsLAwMC/v7++vr69vb28vLy7u7u5
    ubm4uLi3t7e2tra1tbXJycnIyMjHx8fBwcG6urrMzMzLy8vKysqRpb/Pz8/Ozs7Nzc2otMVgjLzS
    0tLR0dHQ0NC8wst3l7+Am7/Bw8fV1dXU1NTT09OesMiXqsOut8Z9m8HHyMrY2NjX19fW1tasucuv
    ucl0mMK9wsp/nMHb29va2trZ2dnCyNCEoMSuusqotsh2msPe3t7d3d3c3Nza2tnQ09aQqMesusuy
    vc2En8LT0dHYu7Xdl4bSt7Lh4eHg4ODf39/d2djdv7nQvb6drMWrmq7amo7ekn/WuLLk5OTj4+Pi
    4uLexcHds6vVy8nDxs3Dxc3LzdHn5+fm5ubCxtGsssacpsCJlrhhdqpYcKdYcaeiqb7q6urp6em2
    vMy5wtWmsMrJz962wNTDytq1vdKrssTt7e3s7Oy7wc+SoL+ap8Olrsestcylr8irtMuyuMfv7+/u
    7u7o6Ojl5eXy8vLw8PAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
    AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
    AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
    AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
    AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
    AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
    AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAItgArCFRQ
    gaCCgwgRFhxY0GDCgwIFAsAhxmBDhRErnOmwUMGEhD4i4hAhQk3ECREVhAwJRQyHMTgAVADwsaCP
    IE08DMHRYWTGiD6CehiphoOInwKDBg0iZMgQDB6cTEkCNGgTJ0/EREEytUKQiFb8KMWSJUsFHDgq
    NFnZJ6wVHEPEjDFTpoyYlRX69PHRR40aNm3UBD6jlK9eK2rGQGEzBMoQNYX1Stbrx21hw5Pb9vHj
    Z+9eHwEBADs=
'''

# Cargo Headers
Param= A5("CdICfg.xlsx","Parametros","ParName")

def Pmt(VarName):
    return Param.D["ParName"][VarName]["ParVal"] 

UpdateChk(Pmt("Update"),".*")

Hoy= date.today()
Fecha= Hoy.strftime("%d/%m/%Y")

C2H= {} # Numero de columna --> Nombre columna
H2C= {} # Nombre columna --> Numero de columna
Lista={} # Para cada columna, los valores posibles
nc= 0
LxF=0
for c in Pmt("DesCol").split("|"):
    C2H[nc]= c
    H2C[c]= nc
    Lista[c]={""}
    LxF+= len(c)
    nc+= 1

Modif= Pmt("Modif").split("|")
kPos= H2C[Pmt("Serial")]
IxL= int(Pmt("ItemsXlinea"))
LxF= str(int(LxF*IxL/3.2))

aInv= A5xlsX(Pmt("PathName")+".xlsm", "Inventario", NoComma0=True)
iKeys= {} # Directorio con clave serial, y dato: una lista con los valores de cada columna separados por ;
kList= {" "}
SerialDuplicado= []
l= 0
for lin in aInv:
    Col= aInv[lin]
    k= Col[kPos].upper()
    if len(Col)==nc:
        if k in iKeys: SerialDuplicado.append(k)
        else:
            Col.append(l)
            blk= ""
            for m in aInv[lin]:
                if m=="None": blk+=";"
                else: blk+=str(m)+";"
            iKeys[k]= blk[0:-1]
            kList.update({k})
            x=0
            for n in range(0,nc):
                if Col[n] not in Lista[C2H[x]]: 
                    Lista[C2H[x]].update({Col[n]})  
                x+= 1
    else: win32ui.MessageBox("Serial ("+k+") con cantidad de columnas "+str(len(Col))+" incorrectas")
    l+= 1

if SerialDuplicado != []:
    win32ui.MessageBox("Los siguientes numeros de serie:\n\n\t"+"\n\t".join(SerialDuplicado)+"\n\nEstan repetidos en el inventario. Este software solo traera el primer registro con esa clave.", "CUIDADO!!!")
# Carga de ubicaciones (grupo , area etc)
NAEG={}
pNae=[]
nNae=0
lNae=0
kNae=-1
with open(Pmt("NAEG")+".csv",errors="backslashreplace") as a:
    for lin in a:
        Col= lin[0:-1].split(";")
        if lNae==0: 
            for c in Col: 
                pNae.append(c)
                if c==Pmt("kUbica"):kNae= lNae 
            nNae= len(Col)
        else:
            if Col[0] not in NAEG: NAEG[Col[0]]={}
            for c in range(0,nNae): NAEG[Col[0]][pNae[c]]= Col[c]
            if Col[kNae] not in Lista[Pmt("Custodio")]: # Agrego Cust. no existentes en inv, pero si en NAEG 
                Lista[Pmt("Custodio")].update({Col[kNae]})
        lNae+=1
    a.close()

kUbica=H2C[Pmt("Custodio")]

Cambios= {} # Clave serial, datos: string completo de datos
# Si ya existian cambios, los leo
if os.path.isfile(Pmt("Cambios")):
    with open(Pmt("Cambios"),errors="backslashreplace") as fp:
        for tx in fp:
            lDat=tx[:-1].split("\t")
            Cambios[lDat[0]]=lDat[1]
    fp.close()


def Cambios_Save(event=" "):
    if os.path.isfile(Pmt("Cambios")): os.remove(Pmt("Cambios"))
    Sale= open(Pmt("Cambios"),"w",newline="\n")
    for s in Cambios:
        Sale.write(str(s)+"\t"+Cambios[s]+"\n")
    Sale.close()


def L2S(Lista,Sep):
    RetVal=""
    for i in Lista: RetVal+=str(i)+Sep
    RetVal= RetVal[:-1]
    return RetVal


def ShowAF(event=" "):
    global Cambios
    if (gui.GetVal("Tecnico") in Lista[Pmt("TecCambio")]) and (len(gui.GetVal("Tecnico"))>1):
        gui.On("Tecnico", False)
        if gui.ValOk("Dato"):
            k= gui.GetVal("Dato")
            if k in iKeys:
                aFrame= Toplevel(mFrame)
                aFrame.geometry(LxF+'x'+str(IxL*15))
                af= A5Tk(aFrame,"Datos del serial "+gui.GetVal("Dato"),Icon=EcoLogo)
                Col= iKeys[k].split(";")
                Primero=""
                fc=0
                fr=0
                for i in Modif:
                    af.Create(i, "e", (fr*2)+1, fc, i, Values=Lista[i],ExChk=r"^[^;\n\t]*$",Horiz=False)
                    af.SetVal(i, Col[H2C[i]])
                    if Primero=="": Primero= i
                    fc+= 1
                    if fc==IxL:
                        fr+=1
                        fc=0
                def Cambio(event=None):
                    # Aplico cambios realizados.
                    for i in Modif: Col[H2C[i]]= re.sub("^\s*|\s*$", "", af.GetVal(i))
                    Nuevo=";".join(Col)
                    if iKeys[k] != Nuevo: 
                        Cambios[k]= Nuevo
                        Cambios_Save()
                        gui.SetVal("Envio", "Enviar "+str(len(Cambios))+" Cambios <E>")
                        gui.On("Envio", True)
                    aFrame.destroy()
                af.Create("Cambio","b",(fr*2)+2, fc,Text="Aceptar Cambios",Values="c",fBind=Cambio)
                af.SetFocus(Primero,True)
        else:
            messagebox.showwarning("Aviso", "Debe seleccionar un dato valido")
            gui.SetFocus("Dato",True)
    else: 
        messagebox.showwarning("Aviso", "Seleccione un Tecnico valido")
        gui.SetFocus("Tecnico",True)


def Envio(event=" "):
    global Cambios
    global nc
    global gui
    if len(Cambios)>0:
        gui.SetVal("Envio", "Sin Cambios")
        gui.On("Envio",False)
        fName=tempfile.gettempdir()+"\\Cambio_Inventario.xlsx"
        Sale=A5(fName,"Inventario",Create=True)
        Sale.SetCell(1, 1, "Movimiento")
        Heads= Pmt("DesCol").split("|")
        for f in range(0,Pmt("Ultimos")): Heads.pop() 
        Sale.SetCell(1, 2, Heads)
        f=2
        a=0 # cantidad de altas
        Items=0
        for k in sorted(Cambios):
            df= 0
            Movi= "Nuevo Activo"
            if k in iKeys:
                Ori= iKeys[k].split(";")
                Sale.SetCell(f,1,"Activo Actual")
                Movi= "Cambio Efectuado"
                a+=1
                df= 1
            New= Cambios[k].split(";")
            Sale.SetCell(f+df,1,Movi)
            uOk= (New[kUbica] in NAEG)
            for c in range(0,nc-Pmt("Ultimos")):
                if df==1: Sale.SetCell(f  , c+2, Ori[c])
                C=C2H[c]
                if C==Pmt("fCambio"):
                    Sale.SetCell(f+df, c+2, Fecha)
                elif C==Pmt("TecCambio"):
                    Sale.SetCell(f+df, c+2, gui.GetVal("Tecnico"))
                elif uOk and (C in NAEG[New[kUbica]]):
                    Sale.SetCell(f+df, c+2, NAEG[New[kUbica]][C])
                else:
                    Sale.SetCell(f+df, c+2, New[c])
            Items+= 1
            f+= 1+df
        Sale.Background("cyan", 1, 1, 1,Title=True)  
        Sale.Save()
        Cambios={}
        # Envio correo
        olMailItem = 0x0
        obj= win32com.client.Dispatch("Outlook.Application")
        mText= "Cambio_Inventario de"+os.getlogin()+" - altas "+str(Items-a)+" sobre total de "+str(Items)
        Correo = obj.CreateItem(olMailItem)
        Correo.To= Pmt("Correo")
        #Correo.CC= "Operaciones"
        Correo.Subject= mText
        Correo.Attachments.Add(Source=fName)
        Correo.BodyFormat= 2
        Correo.HTMLBody= mText
        Correo.Display(True)
        if os.path.isfile(Pmt("Cambios")): os.remove(Pmt("Cambios"))
    else:
        messagebox.showwarning("Aviso", "No aplico cambios para enviar")


def Baja(event=""):
    bFrame= Toplevel(mFrame)
    bFrame.wm_state('iconic')
    bFrame.geometry("+100+250")
    Bajas=A5Tk(bFrame,"Baja de activos",Icon=EcoLogo)
    
    def GeneraBaja(event=""):
        srlist=re.sub("(\n\s*)*$","", Bajas.GetVal("Serial")).upper()
        srlist=srlist.split("\n")
        srBajas=[]
        for s in range(0,len(srlist)): 
            if srlist[s]!="": srBajas.append(srlist[s]) 
        if len(srBajas)==0: messagebox.showwarning("Aviso", "Encuentro dificil dar de baja cero activos!...")
        else:
            Sale=""
            ListaBajas=[]
            for s in srBajas:
                if s in iKeys:
                    campo=iKeys[s].split(";")
                    if campo[H2C[Pmt("mBaja")]]!="": Sale+=s+" "
                    else:ListaBajas.append(s)
                else: Sale+=s.lower()+" "
            if Sale!="": messagebox.showwarning("Aviso", "Los siguientes seriales no estan activos en el inventario\n"+Sale)
            else:
                global Cambios
                Plantilla=[]
                mb=Bajas.GetVal("Razon")
                fb=Bajas.GetVal("Fecha")
                cb=Bajas.GetVal("RadBaja")
                n=0
                t="\t"
                for s in ListaBajas:
                    n+=1
                    campo=iKeys[s].split(";")
                    if cb: 
                        campo[H2C[Pmt("eAF")]]= ""
                        campo[H2C[Pmt("mBaja")]]= mb
                        campo[H2C[Pmt("fBaja")]]= fb
                    else:
                        campo[H2C[Pmt("eAF")]]= "Para Baja"
                    Cambios[s]=";".join(campo)
                    if campo[H2C[Pmt("AF")]]=="":campo[H2C[Pmt("AF")]]="S/N"
                    Plantilla.append(campo[H2C[Pmt("AF")]]+t+"1"+t+campo[H2C[Pmt("dAF")]]+" s/n "+s+t+t+t+t+t+t+t+t+t+t+t+t+t+t+t+t+t+t+mb)
                Bajas.SetVal("Serial", "\n".join(Plantilla))
                Cambios_Save()
                gui.SetVal("Envio", "Enviar "+str(len(Cambios))+" Cambios <E>")
                gui.On("Envio", True)
    
    Bajas.Create("Fecha", "E", 1, 0, "Fecha de baja cont. (DD/MM/AAAA)", Values=Fecha,Horiz=False, ExChk=r"[0123]\d\/[01]\d\/20\d\d")
    Bajas.Create("Razon", "e", 3, 0, Text="Motivo de baja", Horiz=False, Values=Lista[Pmt("mBaja")], ExChk="[^ ]+")
    Bajas.SetVal("Razon", "Obsolescencia")
    Bajas.Create("Serial", "t", 0, 1,Text="Lista de seriales (1 por linea)", Horiz=False, Span=6)
    Bajas.SetWH("Serial", 80,6)
    Bajas.Create("RadBaja", "c", 6, 0, Text="Baja definitiva")
    Bajas.SetVal("RadBaja", True)
    Bajas.Create("Baja", "B",6,1, Text="Generar",Values="g", fBind=GeneraBaja)
    Bajas.SetFocus("Fecha", True)


def ShowCust(event=" "):
    if (gui.GetVal("Tecnico") in Lista[Pmt("TecCambio")]) and (len(gui.GetVal("Tecnico"))>1):
        gui.On("Tecnico", False)
        if gui.ValOk("Dato"):
            CUST=gui.GetVal("Dato")
            cFrame= Toplevel(mFrame)
            cFrame.wm_state('iconic')
            cFrame.geometry("+100+250")
            cu= A5Tk(cFrame,"A.F. que posee el custodio "+CUST,Icon=EcoLogo)
            
            def Swap(event=" "):
                global Cambios
                srlist=re.sub("(\n\s*)*$","", cu.GetVal("swSerial")).upper()
                newSerial= srlist.split("\n")
                oldSerial= cu.GetChkBoxs()
                if len(newSerial) == len(oldSerial):
                    sw=0
                    for ns in oldSerial: # Para cada serial a rotar
                        if newSerial[sw] in iKeys:
                            New=iKeys[newSerial[sw]].split(";")
                            Ori=iKeys[ns].split(";")
                            # Intercambio datos
                            for h in Pmt("DesCol").split("|"):
                                if h in Pmt("Swap"):
                                    c=H2C[h]
                                    Ori[c], New[c] = New[c], Ori[c]
                            Cambios[newSerial[sw]]=";".join(Ori)  
                            Cambios[ns]=";".join(New)
                            cu.SetVal(ns, False)
                            sw+=1
                        else: messagebox.showwarning("Aviso", "El serial "+newSerial[sw]+" No existe en el inventario.\nCorrijalo y reintente")
                    if not cu.GetChkBoxs():
                        gui.On("Envio",True)
                        gui.SetVal("Envio", "Enviar "+str(len(Cambios))+" Cambios <E>")
                        gui.SetVal("Dato", "")
                        Cambios_Save()
                        cFrame.destroy()
                else:messagebox.showwarning("Aviso", "Debe poner "+str(len(oldSerial))+" seriales para poder\nrotarlos con los que selecciono")
            
            nl= 0
            nc= 0
            cu.Create("Serial", "l", nl, nc, "Serial")
            for h in Pmt("ChCust").split("|"): 
                nc+= 1
                cu.Create(h, "l", nl, nc, h,Span=1)
            # Muestro seriales y datos del Custodio
            Primero=""
            for s in sorted(kList):
                if (s!=" ") and (nl<29):
                    c=iKeys[s].split(";")
                    if CUST == c[H2C[Pmt("Custodio")]]:
                        nl+= 1
                        cu.Create(s, "c", nl, 0, s,Span=1) # Serial encontrado falta fbind
                        if Primero=="": Primero=s
                        # Muestro datos del serial encontrado.
                        nc= 0
                        for h in Pmt("ChCust").split("|"):
                            nc+= 1
                            cu.Create("caso"+str(nl)+str(nc), "l", nl, nc, Text=c[H2C[h]],Span=1)
            if nl==0: 
                cFrame.destroy()
                messagebox.showwarning("Aviso", "Custodio sin activos asignados")
            else:
                def AddSrl(event=" "):
                    txt=cu.GetVal("swSerial")
                    nk= cu.GetVal("prusrl")
                    if nk in iKeys:
                        cu.SetVal("swSerial", txt+nk+"\n")
                        txt=cu.GetVal("serTipo")
                        sTipo= iKeys[nk].split(";")
                        cu.SetVal("serTipo", txt+sTipo[H2C[Pmt("Tipo")]]+"\n")
                        cu.SetVal("prusrl","")
                def bserial(event=" "):
                    cu.SetFocus("prusrl", True)
                cu.Create("serTipo", "t",nl+1, 0, Text="Tipo del serial a la derecha ->",Horiz=False)
                cu.SetWH("serTipo", 20, 5)
                cu.Create("swSerial", "t",nl+1,1, Text="Intercambio (en orden,1 por linea)",Horiz=False)
                cu.SetWH("swSerial", 20, 5)
                cu.Create("prusrl", "e", nl+1, 2, "Buscar <S>erial", kList, fBind=AddSrl)
                cu.Create("swap", "b", nl+2, 6, Text="Rotar seriales",Values="r",fBind=Swap,Span=nc)
                cFrame.bind("<Alt_L><s>",bserial)
                cu.SetFocus(Primero,True)
        else:
            messagebox.showwarning("Aviso", "Debe seleccionar un dato valido")
            gui.SetFocus("Dato",True)


def Nuevos(event=""):
    if (gui.GetVal("Tecnico") in Lista[Pmt("TecCambio")]) and (len(gui.GetVal("Tecnico"))>1):
        gui.On("Tecnico", False)
        aFrame= Toplevel(mFrame)
        af= A5Tk(aFrame,"Alta de Activos Fijos",Icon=EcoLogo)
        af.Create("afs", "t", 0, 0, "Numeros de AF", Horiz=False)
        af.Create("srl", "t", 0, 1, "Nuevos Seriales", Horiz=False)
        n=af.GetObj("afs").configure(height=4,width=20)
        af.GetObj("srl").configure(height=4,width=20)
        Primero=""
        for i in Pmt("Alta").split("|"):
            af.Create(i, "e", Text=i, Values=Lista[i],ExChk=r"^[^;\n\t]*$")
            if Primero=="": Primero= i
        af.SetFocus(i)
        def Genera(event=""):
            def limpio(lStr):
                s= af.GetVal(lStr)
                s= re.sub(r"\n\s*\n", "\n", s)
                s= re.sub(r"\n\s*$", "", s)
                return s.split("\n")
            NroAFs=limpio("afs")
            NroSrl=limpio("srl")
            AF= H2C[Pmt("AF")]
            SR= H2C[Pmt("Serial")]
            if len(NroAFs)==len(NroSrl):
                Col=[]
                for i in range(1,nc):Col.append("")
                for s in range(0,len(NroSrl)):
                    for i in Pmt("Alta").split("|"): Col[H2C[i]]= re.sub("^\s*|\s*$", "", af.GetVal(i)) 
                    Col[AF]= NroAFs[s]
                    Col[SR]= NroSrl[s]
                    Cambios[NroSrl[s]]=";".join(Col)
                gui.On("Envio",True)
                gui.SetVal("Envio", "Enviar "+str(len(Cambios))+" Cambios <E>")
                Cambios_Save()
                aFrame.destroy()
            else:messagebox.showwarning("Aviso", "Las cantidades de Activos Fijos y Seriales, deben ser iguales.")
        af.Create("Genera", "b", Text="Generar altas",Values="g", fBind=Genera)
        af.SetFocus("afs")
    else: 
        messagebox.showwarning("Aviso", "Seleccione un Tecnico valido")
        gui.SetFocus("Tecnico",True)
    

def CustSer(event=" "):
    if gui.GetVal("ChkSerial"):
        gui.Set("Dato", "Values", kList)
        gui.Set("Dato", "ExChk", r"^[\dA-Za-z_]+$")
        gui.SetSimil("Dato", Sets=True)
        gui.Set("Dato", "fBind", ShowAF)
    else:
        gui.Set("Dato", "Values", Lista[Pmt("Custodio")])
        gui.Set("Dato", "ExChk", r"^\d{5} [^\t;]+$")
        gui.SetSimil("Dato", Sets=False)
        gui.Set("Dato", "fBind", ShowCust)


mFrame= Tk()
loginname=os.getlogin()[2:]
gui= A5Tk(mFrame, "Act. Fijos. v1.3.04", TopLeft=[100,100], Icon=EcoLogo)
gui.Create("Tecnico", "e", 0, 0, "Tecnico: ",Values=Lista[Pmt("TecCambio")])
gui.GetObj("Tecnico").configure(width=30)
gui.SetVal("Tecnico", loginname)
gui.Create("ChkSerial", "c", 1, 1, Text="Seleccionar por Serial", fBind=CustSer,Span=1)
gui.SetVal("ChkSerial", True)
gui.Create("Dato", "e", 2, 0, "Ingrese dato")
gui.GetObj("Dato").configure(width=30)
CustSer()
gui.Create("Envio", "b", 3, 1, "Sin cambios",Values="e", fBind=Envio)
gui.GetObj("Envio").configure(width=25)
gui.Create("Nuevo", "b", 3, 0, "Nuevos AF",Values="n", fBind=Nuevos)
if len(Cambios)>0: gui.SetVal("Envio","Enviar "+str(len(Cambios))+" Cambios")
else: gui.On("Envio", False)
mFrame.bind("<Alt_L><b>",Baja)
gui.SetFocus("Tecnico")
mFrame.mainloop()

