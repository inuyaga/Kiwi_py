from tkinter import *
from tkinter.ttk import *
import os   
import productos
from tkinter import messagebox, simpledialog
from tkinter import filedialog

from xlwt import Workbook
from xlwt import Font
from xlwt import XFStyle
from xlwt import Borders
class Kiwi:
    def __init__(self):
        self.rootKiwi=Tk()
        # bit = root.iconbitmap('MiImagen.ico')
        # self.rootKiwi.iconbitmap(r'/home/esteban/Documentos/python/Kiwi/kiwi.ico')
        imgicon = PhotoImage(file=os.path.join('kiwi.gif')) 
        self.rootKiwi.tk.call('wm', 'iconphoto', self.rootKiwi._w, imgicon) 
        # self.rootKiwi.attributes('-zoomed', True) 
        ancho_windows=self.rootKiwi.winfo_screenwidth()
        alto_windows=self.rootKiwi.winfo_screenheight()
        # geometry_data="{}x{}+0+0".format(ancho_windows, alto_windows)
        # self.rootKiwi.geometry(geometry_data)
        self.Productos=productos.Productos()
        self.rootKiwi.title('Kiwi')
        self.MenuBar()
        self.KiwiFrame=Frame(self.rootKiwi)
        self.KiwiFrame.pack(expand=0, fill=X)
        self.elementUI()
        self.rootKiwi.mainloop()

    def MenuBar(self):
        menubar = Menu(self.rootKiwi)
        filemenu = Menu(menubar, tearoff=0)
        filemenu.add_command(label="Venta nueva", command=self.clear_venta)
        # filemenu.add_command(label="Venta nueva", command=self.clear_venta)

        filemenu.add_separator()
        filemenu.add_command(label="Salir", command=self.salir)
        

        ayudamenu=Menu(menubar, tearoff=0)
        ayudamenu.add_command(label="Acerca de...", command=self.acercade)


        menubar.add_cascade(label="ARCHIVO", menu=filemenu)
        menubar.add_cascade(label="AYUDA", menu=ayudamenu)
        self.rootKiwi.config(menu=menubar)

    
    def acercade(self):
        msn = """ 
        Este software ha sido desarrolado por ISC. Mario Esteban Hernandez Hernandez ® derechos reservados..
        """
        messagebox.showinfo(message=msn, title="Acerca de")
    
    def clear_venta(self):
        self.tree2.delete(*self.tree2.get_children())
    def salir(self):
        self.rootKiwi.destroy()

    def elementUI(self):
        # CREAMOS LABEL Y CAJA DE TEXTO
        Label(self.KiwiFrame, text='Buscar:').grid(row=0, column=0, sticky='w')
        Label(self.KiwiFrame, text='PRECIOS SUGERIDOS:').grid(row=0, column=3)
        self.varBuscador=StringVar()
        self.txt_buscardor=Entry(self.KiwiFrame, textvariable=self.varBuscador, text='Codigo producto')
        self.txt_buscardor.grid(row=0, column=1, padx=10, pady=10, columnspan=2, sticky='nsew')
        # ASIGNAMOS UNA ACCION A LA CAJA DE TEXTO
        self.txt_buscardor.bind("<Return>", self.key)

        self.add_item=Button(self.KiwiFrame, text="Agregar producto", width=7, command=lambda:self.add_producto())
        self.add_item.grid(row=11, column=3, sticky="ew")

        # CREAMOS UN TREEVIEW Y SUS COLUMNAS
        self.tree = Treeview(self.KiwiFrame, columns=('description'))
        self.col1=self.tree.heading('#0', text='##')
        self.tree.heading('description', text='DESCRIPCIÓN')
        # CONFIGURAR TAMAÑO DE COLUMNAS
        self.tree.column('#0', width=150, anchor='center')
        self.tree.column('description', width=500, anchor='w')

        # AGREGAMOS TABLA AL FRAME
        self.tree.grid(row=1, column=1, rowspan=10, sticky='nsew')
        vsb = Scrollbar(self.KiwiFrame, orient="vertical", command=self.tree.yview)
        vsb.grid(row=1, column=2, sticky='nsew', rowspan=10)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.bind("<<TreeviewSelect>>", self.TreeSelectedItem)

        self.varOptions=IntVar()

        self.R1=Radiobutton(self.KiwiFrame,variable=self.varOptions, text='PRECIO 1:', value=1, command=self.get_var_option)
        self.R2=Radiobutton(self.KiwiFrame,variable=self.varOptions, text='PRECIO 2:', value=2, command=self.get_var_option)
        self.R3=Radiobutton(self.KiwiFrame,variable=self.varOptions, text='PRECIO 3:', value=3, command=self.get_var_option)
        self.R4=Radiobutton(self.KiwiFrame,variable=self.varOptions, text='PRECIO 4:', value=4, command=self.get_var_option)
        self.R5=Radiobutton(self.KiwiFrame,variable=self.varOptions, text='PRECIO 5:', value=5, command=self.get_var_option)
        self.R6=Radiobutton(self.KiwiFrame,variable=self.varOptions, text='PRECIO 6:', value=6, command=self.get_var_option)
        self.R7=Radiobutton(self.KiwiFrame,variable=self.varOptions, text='PRECIO 7:', value=7, command=self.get_var_option)
        self.R8=Radiobutton(self.KiwiFrame,variable=self.varOptions, text='PRECIO 8:', value=8, command=self.get_var_option)
        self.R9=Radiobutton(self.KiwiFrame,variable=self.varOptions, text='PRECIO 9:', value=9, command=self.get_var_option)
        self.R10=Radiobutton(self.KiwiFrame,variable=self.varOptions, text="PRECIO 10:", value=10, command=self.get_var_option)

        self.R1.grid(row=1, column=3, columnspan=2)
        self.R2.grid(row=2, column=3, columnspan=2)
        self.R3.grid(row=3, column=3, columnspan=2)
        self.R4.grid(row=4, column=3, columnspan=2)
        self.R5.grid(row=5, column=3, columnspan=2)
        self.R6.grid(row=6, column=3, columnspan=2)
        self.R7.grid(row=7, column=3, columnspan=2)
        self.R8.grid(row=8, column=3, columnspan=2)
        self.R9.grid(row=9, column=3, columnspan=2)
        self.R10.grid(row=10, column=3, columnspan=2)


        # CREAMOS LA SEGUNDA TABLA DE VENTA
        self.tree2 = Treeview(self.KiwiFrame, columns=('description', 'precio', 'cantidad', 'total'))
        # AÑADIMOS LOS ENCABEZADOS
        self.tree2.heading('#0', text='##')
        self.tree2.heading('description', text='DESCRIPCION')
        self.tree2.heading('precio', text='PRECIO')
        self.tree2.heading('cantidad', text='CANTIDAD')
        self.tree2.heading('total', text='TOTAL')
        # CONFIGURAMOS TAMAÑO TABLA
        self.tree2.column('#0', width=70, anchor='w')
        self.tree2.column('description', width=200, anchor='w')
        self.tree2.column('precio', width=50, anchor='center')
        self.tree2.column('cantidad', width=50, anchor='center')
        self.tree2.column('total', width=50, anchor='center')

        # agregamos un scroll a la tabla 2 
        vsb2 = Scrollbar(self.KiwiFrame, orient="vertical", command=self.tree2.yview)
        vsb2.grid(row=12, column=2, sticky='nsew')

        # AGREGAMOS UN LABEL A LA TABLA
        Label(self.KiwiFrame, text='VENTA ACTUAL').grid(row=11, column=0, columnspan=3)

        # menu contextual a la tabla 2
        self.aMenu = Menu(self.KiwiFrame, tearoff=0)
        self.aMenu.add_command(label='Elimninar', command=self.DeleteItemTree)
        self.tree2.bind("<Button-3>", self.popup)
        # self.aMenu.add_command(label='Say Hello', command=self.hello)

        # agregamos la tabla 2 al frame 
        self.tree2.grid(row=12, column=1, sticky='nsew')

        self.genera_xls=Button(self.KiwiFrame, text="Generar XLS", width=7, command=lambda:self.crear_xls())
        self.genera_xls.grid(row=13, column=3, columnspan=2,  sticky='nsew')
    
    def DeleteItemTree(self):
        respuesta=messagebox.askyesno(message="Eliminar producto de venta actual, ¿Desea continuar?", title="Eliminar producto")
        if respuesta:
            selected_item = self.tree2.selection()[0] ## get selected item
            self.tree2.delete(selected_item)  
    def popup(self, event):
        self.iid = self.tree2.identify_row(event.y)
        if self.iid:
            # mouse pointer over item
            self.tree2.selection_set(self.iid)
            self.aMenu.post(event.x_root, event.y_root)          
        else:
            pass

        
    def add_producto(self):
        item=self.tree.selection()
        if len(item) > 0:
            if self.varOptions.get() > 0:
                # MOSTRAMOS CUADRO DE DIALOGO PARA INGRESAR Cantidad
                number_prod = simpledialog.askstring(title="Cantidad", prompt="Numero de productos")
                # OBTENEMOS EL VALOR EN LA SELECCION DE LA TABLA EN ESTE CASO CODIGO PRODUCTO
                item=self.tree.selection()[0]
                codigo_producto=self.tree.item(item, option="text")
                # OBTENEMOS EL VALOR DEL RADIO BUTTON EN EL MOMENTO
                precio_radio_selected=self.varOptions.get()

                # ESTABLECEMOS UNA CONDICION DE INGRESO DE DATOS QUE SOLO SEAN ENTEROS
                if number_prod != None and number_prod != '':
                    try:
                        cantidad_pedir = int(number_prod)
                        datos=(codigo_producto,)
                        query_datos=self.Productos.get(datos)
                        txt_precio = "PRECIO{}".format(precio_radio_selected)
                        sub_total=float(query_datos[txt_precio]) * float(cantidad_pedir)
                        self.tree2.insert("", END, text=query_datos['PRODUCTO'],values=(
                            query_datos['DESC1'],
                            query_datos[txt_precio],
                            cantidad_pedir,
                            sub_total,
                        ))
                    except ValueError as identifier:
                        messagebox.showerror("Incorrecto", "Ingrese solo numeros enteros")                
            else:
                messagebox.showinfo("Elegir precio", "Antes debe de seleccionar algun precio sugerido")
        else:
            messagebox.showinfo("Ningun producto seleccionado", "Debe seleccionar un elemento o buscar un producto antes")
            
        
        # var=self.tree.item(item, option="text")
        
    
    def TreeSelectedItem(self, event):
        # self.selected = event.widget.selection()
        # for idx in self.selected:
        #     print(self.tree.item(idx)['text'])
        item=self.tree.selection()[0]
        var=self.tree.item(item, option="text")
        datos=(var,)
        query_datos=self.Productos.get(datos)
        # print(query_datos['PRECIO1'])
        self.R1['text']="PRECIO 1: ${}".format(query_datos['PRECIO1'])
        self.R2['text']="PRECIO 2: ${}".format(query_datos['PRECIO2'])
        self.R3['text']="PRECIO 3: ${}".format(query_datos['PRECIO3'])
        self.R4['text']="PRECIO 4: ${}".format(query_datos['PRECIO4'])
        self.R5['text']="PRECIO 5: ${}".format(query_datos['PRECIO5'])
        self.R6['text']="PRECIO 6: ${}".format(query_datos['PRECIO6'])
        self.R7['text']="PRECIO 7: ${}".format(query_datos['PRECIO7'])
        self.R8['text']="PRECIO 8: ${}".format(query_datos['PRECIO8'])
        self.R9['text']="PRECIO 9: ${}".format(query_datos['PRECIO9'])
        self.R10['text']="PRECIO 10: ${}".format(query_datos['PRECIO10'])


    def get_var_option(self):
        pass

    def Tree2Action(self, event):
        pass
        


    def crear_xls(self):
        # item=self.tree.selection()[0]
        # var=self.tree.item(item, option="text")
        # print(var)
        # Workbook asing


     

        self.filename =  filedialog.asksaveasfilename(initialdir = "/",title = "Guardar venta",filetypes = (("Archivo xls","*.xls"),("all files","*.*")))
        if len(self.filename) > 0:
            first_book=Workbook()

            # Sheets definition
            ws1=first_book.add_sheet('venta')

            header_font = Font()
            body_font = Font()
            # Header font preferences
            header_font.name = 'Times New Roman'
            header_font.height = 20 * 10
            header_font.bold = True
            # Body font preferences
            body_font.name = 'Arial'
            body_font.italic = True
            # Header Cells style definition
            header_style = XFStyle()
            header_style.font = header_font 
            borders = Borders()
            borders.left = 1
            borders.right = 1
            borders.top = 1
            borders.bottom = 1
            header_style.borders = borders
            # body cell name style definition
            body_style = XFStyle()
            body_style.font = body_font 

            print(ws1.col(0).width)
            ws1.write(0, 0, '##', header_style)
            ws1.write(0, 1, 'DESCRIPCION', header_style)
            ws1.write(0, 2, 'PRECIO', header_style)
            ws1.write(0, 3, 'CANTIDAD', header_style)
            ws1.write(0, 4, 'SUBTOTAL', header_style)
            ws1.col(0).width = 2962 * 2
            ws1.col(1).width = 2962 * 4
            ws1.col(2).width = 2962 * 1
            ws1.col(3).width = 2962 * 1
            ws1.col(4).width = 2962 * 1

            tabla_venta=self.tree2.get_children()
            for item, conteo in zip(tabla_venta, range(len(tabla_venta))):
                ws1.write(conteo+1, 0, self.tree2.item(item)['text'])

                ws1.write(conteo+1, 1, self.tree2.item(item)['values'][0], body_style)
                ws1.write(conteo+1, 2, self.tree2.item(item)['values'][1], body_style)
                ws1.write(conteo+1, 3, self.tree2.item(item)['values'][2], body_style)
                ws1.write(conteo+1, 4, self.tree2.item(item)['values'][3], body_style)

                # print(self.tree2.item(item)["values"])
                # print(self.tree2.item(item)['text'])

            

            # Saving file
            first_book.save(self.filename)


        

    def key(self, event):
        self.tree.delete(*self.tree.get_children())
        datos=(self.txt_buscardor.get(), self.txt_buscardor.get())
        query_datos=self.Productos.consulta(datos)
        if len(query_datos) > 0:
            for item in query_datos:
                self.tree.insert("", END, text=item['PRODUCTO'],values=(item['DESC1'],))
        else:
            messagebox.showinfo("Sin resultados", "Producto no encontrado")
        
        
            
        
        

api=Kiwi()