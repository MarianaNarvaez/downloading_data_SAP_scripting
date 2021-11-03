systemP15: str = r"1.01  (P15) - Global ERP Production"
systemP11: str = r"1.02  (P11) - European ERP Production"
systemRPD: str = r"1.13  (RPD) - MBPS ECC 6.0 - Production"

userP15 = ""
pwdP15 = r""

userP11 = ""
pwdP11 = r""

userRPD = ""
pwdRPD = ""

path_data_local: str = r"D:\Files\ME"
path_data_local_rpd: str = r"D:\Files\ME\RPD"

clase_coste_rpd = '\r\n'.join(['51140001', '51070017', '51150002', '51150001',
                               '51150003', '51150004', '51150005', '51150006', '51150007'])

p11_me_files = [path_data_local + r'\p11_ksb1_me_I35600.xlsx',
                path_data_local + r'\p11_ksb1_me_I35300.xlsx',
                path_data_local + r'\p11_ksb1_me_MEREG.xlsx']

CONN = r"""Driver={SQL Server};Server="";Database="";UID="";PWD="""""

database_p = r""
userandpass = r''


COLUMNS_P11_KSB1 = ['Company_code', 'Plant', 'Cost_center', 'CO_object_name', 'Cost_element', 'Cost_element_name', 'Cost_element_descr', 'Offsetting_account_type', 'Offsetting_acct_no', 'Name_of_offsetting_account', 'Name_of_offsetting_account1', 'Document_type', 'Purchasing_document', 'Posting_date', 'Fiscal_year', 'Period', 'Material', 'Material_description', 'Purchase_order_text', 'Name',
                    'Posted_unit_of_meas', 'Total_quantity', 'Object_currency', 'Value_in_obj_crcy', 'Transaction_currency', 'Value_trancurr', 'Report_currency', 'Val_in_rep_cur', 'Cost_element_group', 'Cost_element_group_name', 'Document_header_text', 'Document_date', 'Document_number', 'Functional_area', 'Material_group', 'Material_group_desc', 'Vendor', 'Vendor_name', 'Value_type', 'Aux_acct_assig','Source', 'Load_date']

COLUMNS_P15_KOC2 = ['Company_code', 'Plant', 'Cost_element', 'Cost_element_name', 'Cost_element_descr', 'Offsetting_account_type', 'Offsetting_acct_no', 'Name_of_offsetting_account', 'Name_of_offsetting_account1', 'Document_type', 'Purchasing_document', 'Posting_date', 'Fiscal_year', 'Period', 'Material', 'Material_description', 'Purchase_order_text', 'Name', 'Posted_unit_of_meas', 'Total_quantity',
                    'Object_currency', 'Value_in_obj_crcy', 'Transaction_currency', 'Value_trancurr', 'Report_currency', 'Val_in_rep_cur', 'Cost_element_group', 'Cost_element_group_name', 'Document_header_text', 'Document_date', 'Document_number', 'Material_group', 'Material_group_desc', 'Value_type', 'Partner_cctr', 'CO_partner_object_name', 'Partner_func_area', 'CO_object_name', 'Functional_area', 'Order', 'Source', 'Load_date']

COLUMNS_P15_KSB1 = ['Company_code', 'Plant', 'Cost_center', 'CO_object_name', 'Cost_element', 'Cost_element_name', 'Cost_element_descr', 'Offsetting_account_type', 'Offsetting_acct_no', 'Name_of_offsetting_account', 'Name_of_offsetting_account1', 'Document_type', 'Purchasing_document', 'Posting_date', 'Fiscal_year', 'Period', 'Material', 'Material_description', 'Purchase_order_text',
                    'Name', 'Posted_unit_of_meas', 'Total_quantity', 'Object_currency', 'Value_in_obj_crcy', 'Transaction_currency', 'Value_trancurr', 'Report_currency', 'Val_in_rep_cur', 'Cost_element_group', 'Cost_element_group_name', 'Document_header_text', 'Document_date', 'Document_number', 'Functional_area', 'Material_group', 'Material_group_desc', 'Value_type', 'Source', 'Load_date']

COLUMNS_RPD_ME2N = ['Documento compras',
                    'Posición',
                    'Cl.documento compras',
                    'Proveedor/Centro suministrador',
                    'Centro',
                    'Material',
                    'Grupo de artículos',
                    'Texto breve',
                    'Fecha documento',
                    'Cantidad de pedido',
                    'Precio neto',
                    'Cantidad base',
                    'Moneda',
                    'Unidad medida pedido',
                    'Historial pedido/Docu.orden entrega',
                    'Organización compras',
                    'Por calcular (cantidad)',
                    'Por calcular (valor)',
                    'Por entregar (valor)',
                    'Por entregar (cantidad)',
                    'Source', 'Load_date']

COLUMNS_RPD_T023T = ['Clave de idioma', 'Grupo_de_articulos', 'Denom_gr_articulos',
                     'Denom_2_gr_articulos', 'Source', 'Load_date']

COLUMNS_RPD_MARA = ['Material', 'Grupo_de_articulos', 'Source', 'Load_date']

COLUMNS_RPD_ME = ['Sociedad', 'Centro', 'Segmento', 'Objeto_del_interlocutor', 'Denom_del_objeto_del_inter', 'Clase_de_coste', 'Denom_clase_de_coste',
                'Descrip_clases_coste', 'Clase_cta_de_contrapartida', 'Cta_contrapartida', 'Denom_cuenta_contrapartida', 
                'Denom_cuenta_contrapartida1', 'Clase_de_documento', 'Documento_compras', 'Fe_contabilizacion', 'Ejercicio', 
                'Periodo', 'Material','Texto_breve_de_material', 'Texto_de_pedido', 'Denominacion', 'Ud_cantidad_contab', 'Cantidad_total_reg', 
                'Moneda_del_objeto', 'Valor_moneda_objeto', 'Moneda_transacción', 'Valor_mon_tr', 'Moneda_sociedad_CO', 'Valor_mon_soc_CO', 
                'Objeto', 'Denom_del_objeto', 'Texto_de_cabecera_de_doc', 'Fecha_de_doc', 'Numero_de_documento', 'Area_funcional', 
                'Tipo_de_valor', 'Orden', 'Posicion']

replaceValP15 = {
    "0" : ['--', ''],
    "1" : ['1\n', ''],
    "2" : ['-\n', ''],
    "3" : ['', ''],
    "4" : ['', '']
}

replaceVal = {
    "0" : ['--', ''],
    "1" : ['-\n', ''],
    "2" : ['', ''],
    "3" : ['', ''],
    "4" : ['', ''],
    "5" : ['', ''],
    "6" : ['', ''],
    "7" : ['', ''],
    "8" : ['', ''],
    "9" : ['', ''],
    "10" : ['', '']
}
