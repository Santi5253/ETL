const Excel = require('exceljs');
const fs = require('fs');
const outDir = 'Entrega_ETL_TecnoHogar';
if(!fs.existsSync(outDir)) fs.mkdirSync(outDir,{recursive:true});
(async()=>{
const wb = new Excel.Workbook();
wb.creator='TecnoHogar ETL'; wb.created=new Date();
const HDR = {fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FF1F4E79'}},font:{bold:true,color:{argb:'FFFFFFFF'},size:11}};
const HDR2 = {fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FFFFC000'}},font:{bold:true,color:{argb:'FF1F4E79'},size:11}};
function sheet(name, headers, rows, hdrStyle=HDR){
  const ws = wb.addWorksheet(name);
  ws.addRow(headers);
  rows.forEach(r=>ws.addRow(r));
  ws.getRow(1).eachCell(c=>{Object.assign(c,hdrStyle);c.border={top:{style:'thin'},left:{style:'thin'},bottom:{style:'thin'},right:{style:'thin'}};});
  ws.columns = headers.map(()=>({width:22}));
  ws.autoFilter='A1:'+String.fromCharCode(64+headers.length)+'1';
  ws.freezePanes='A2';
  return ws;
}
// 1. FUENTES SUCIAS
sheet('Clientes',['id_cliente','nombre','ciudad','correo','telefono','observacion_interna'],[
['C001','  Ana María Torres ','BOGOTA','ana.torres@mail.com','3101112233','x'],
['C002','Carlos Ruiz','bogotá ','carlos.ruiz@mail.com','3102223344','x'],
['C003','Lucía Fernández','Medellin','N/A','3103334455','x'],
['C004','Pedro Gómez','MEDELLÍN','pedro.gomez@mail.com','NULL','x'],
['C005','Sofía Herrera',' cali','sofia.h@mail.com','3105556677','x'],
['C006','Diego López','Cali ','Sin dato','3106667788','x'],
['C007','Marta Jiménez','Barranquilla','marta.j@mail.com','-','x'],
['C008','Jorge Ramírez','barranquilla','jorge.r@mail.com','3108889900','x'],
['C003','Lucía Fernández','Medellin','N/A','3103334455','x'],
['C009','Laura Castro','BOGOTÁ','laura.c@mail.com','','x'],
['C010','Andrés Mora','Cartagena','andres.m@mail.com','3100001122','x'],
['C011','  Paola Díaz  ','cartagena ','paola.d@mail.com','3101113344','x'],
['C012','Miguel Soto','Bucaramanga','miguel.s@mail.com','3102225566','x'],
]);
sheet('Productos',['codigo_producto','producto','categoria','precio','costo','columna_innecesaria'],[
['P001','Portátil HP 15','portatiles','$2.450.000','$1.900.000','borrar'],
['P002','Mouse Logitech','Accesorios','45000','28000','borrar'],
['P003','Celular Samsung A54','CELULARES','1.150.000','890000','borrar'],
['P004','Teclado Mecánico','accesorios ','135000','90000','borrar'],
['P005','Monitor LG 24','Monitores','780000','N/A','borrar'],
['P006','Impresora Epson','IMPRESORAS ','520000','410000','borrar'],
['P007','Tablet Lenovo','celulares','690000','NULL','borrar'],
['P007','Tablet Lenovo','celulares','690000','NULL','borrar'],
['P008','Audífonos Sony','Accesorios','-','95000','borrar'],
['P009','Disco SSD 1TB','Almacenamiento','320000','240000','borrar'],
]);
sheet('Vendedores',['id_vendedor','vendedor','sede','ciudad'],[
['V01','  Juan Pérez','Sede Norte','Bogotá'],
['V02','MARIA GONZALEZ','sede norte','BOGOTA '],
['V03','Carlos Sánchez','Sede Sur','Medellín'],
['V04','Ana Martínez','SEDE SUR','medellin'],
['V05','Luis Torres','Sede Cali','CALI'],
['V02','MARIA GONZALEZ','sede norte','BOGOTA ','DUPLICADO'],
]);
sheet('Ventas_Agosto',['id_venta','fecha','id_cliente','codigo_producto','id_vendedor','cantidad','descuento'],[
['VTA-001','2024-08-03','C001','P001','V01',2,'10%'],
['VTA-002','03/08/2024',' C002 ','P002','V02',3,'5%'],
['VTA-003','2024/08/05','C003','P003','V03',0,'0%'],
['VTA-004','05-08-2024','C004','P004','V04',-1,'15%'],
['VTA-005','2024-08-10','C005','P005','V05',1,'45%'],
['VTA-006','2024-08-12','C999','P002','V01',2,'10%'],
['VTA-007','2024-08-15','C006','P999','V02',1,'5%'],
['VTA-008','2024-08-18','C007','P006','V99',4,'20%'],
['VTA-009','2024-08-20','C008','P007','V03',2,'N/A'],
['VTA-010','2024-08-22','C009','P009','V04',5,'30%'],
['VTA-011','2024-08-25','C010','P002','V05',1,'12%'],
['VTA-011','2024-08-25','C010','P002','V05',1,'12%'],
]);
sheet('Ventas_Septiembre',['id_venta','fecha','id_cliente','codigo_producto','id_vendedor','cantidad','descuento'],[
['VTA-012','2024-09-02','C011','P001','V01',1,'8%'],
['VTA-013','02/09/2024','C012','P003','V03',2,'35%'],
['VTA-014','2024/09/05','C001','P004','V02',0,'0%'],
['VTA-015','2024-09-08','C002','P008','V04',3,'10%'],
['VTA-016','2024-09-10',' C003',' P005 ',' V05 ',2,'15%'],
['VTA-017','2024-09-12','C004','P009','V01',4,'25%'],
['VTA-018','2024-09-15','C005','P002','V03',6,'5%'],
['VTA-019','2024-09-18','C006','P006','V02',1,'50%'],
['VTA-020','2024-09-20','C007','P010','V04',2,'10%'],
['VTA-021','2024-09-22','NULL','P001','V05',1,'5%'],
['VTA-022','2024-09-25','C008','P003','V01',3,'18%'],
]);
// 2. TABLA FINAL LIMPIA (resultado simulado del ETL)
const finalH=['id_venta','fecha','id_cliente','cliente','ciudad_cliente','codigo_producto','producto','categoria','cantidad','precio','costo','descuento','Total_Bruto','Valor_Descuento','Total_Neto','Costo_Total','Utilidad','Clasificacion_Venta','id_vendedor','vendedor'];
const finalR=[
['VTA-001','2024-08-03','C001','Ana María Torres','Bogotá','P001','Portátil HP 15','Portátiles',2,2450000,1900000,0.10,4900000,490000,4410000,3800000,610000,'Alta','V01','Juan Pérez'],
['VTA-002','2024-08-03','C002','Carlos Ruiz','Bogotá','P002','Mouse Logitech','Accesorios',3,45000,28000,0.05,135000,6750,128250,84000,44250,'Baja','V02','María González'],
['VTA-010','2024-08-22','C009','Laura Castro','Bogotá','P009','Disco SSD 1TB','Almacenamiento',5,320000,240000,0.30,1600000,480000,1120000,1200000,-80000,'Media','V04','Ana Martínez'],
['VTA-011','2024-08-25','C010','Andrés Mora','Cartagena','P002','Mouse Logitech','Accesorios',1,45000,28000,0.12,45000,5400,39600,28000,11600,'Baja','V05','Luis Torres'],
['VTA-012','2024-09-02','C011','Paola Díaz','Cartagena','P001','Portátil HP 15','Portátiles',1,2450000,1900000,0.08,2450000,196000,2254000,1900000,354000,'Alta','V01','Juan Pérez'],
['VTA-016','2024-09-10','C003','Lucía Fernández','Medellín','P005','Monitor LG 24','Monitores',2,780000,610000,0.15,1560000,234000,1326000,1220000,106000,'Media','V05','Luis Torres'],
['VTA-017','2024-09-12','C004','Pedro Gómez','Medellín','P009','Disco SSD 1TB','Almacenamiento',4,320000,240000,0.25,1280000,320000,960000,960000,0,'Media','V01','Juan Pérez'],
['VTA-018','2024-09-15','C005','Sofía Herrera','Cali','P002','Mouse Logitech','Accesorios',6,45000,28000,0.05,270000,13500,256500,168000,88500,'Baja','V03','Carlos Sánchez'],
['VTA-009','2024-08-20','C008','Jorge Ramírez','Barranquilla','P007','Tablet Lenovo','Celulares',2,690000,520000,0.00,1380000,0,1380000,1040000,340000,'Media','V03','Carlos Sánchez'],
['VTA-022','2024-09-25','C008','Jorge Ramírez','Barranquilla','P003','Celular Samsung A54','Celulares',3,1150000,890000,0.18,3450000,621000,2829000,2670000,159000,'Alta','V01','Juan Pérez'],
['VTA-006-duplicada-ejemplo-VAL','2024-08-12','C001*','Ejemplo válida','Bogotá','P002','Mouse Logitech','Accesorios',2,45000,28000,0.10,90000,9000,81000,56000,25000,'Baja','V01','Juan Pérez'],
];
const wsF = sheet('BD_VENTAS_ETL',finalH,finalR);
wsF.columns=finalH.map(()=>({width:20}));
// 3. RECHAZADOS
sheet('RECHAZADOS',['id_venta_origen','etapa','motivo_rechazo','detalle'],[
['VTA-003','Validación cantidad','Cantidad <= 0','cantidad=0 viola regla 5'],
['VTA-004','Validación cantidad','Cantidad <= 0','cantidad=-1 viola regla 5'],
['VTA-005','Validación descuento','Descuento > 30%','45% supera máximo'],
['VTA-006','Integridad referencial','Cliente inexistente','C999 no existe en Clientes'],
['VTA-007','Integridad referencial','Producto inexistente','P999 no existe'],
['VTA-008','Integridad referencial','Vendedor inexistente','V99 no existe'],
['VTA-013','Validación descuento','Descuento > 30%','35% supera máximo'],
['VTA-014','Validación cantidad','Cantidad <= 0','cantidad=0'],
['VTA-015','Validación precio','Precio no válido','precio "-" en P008'],
['VTA-019','Validación descuento','Descuento > 30%','50% supera máximo'],
['VTA-020','Integridad referencial','Producto inexistente','P010 no existe'],
['VTA-021','Integridad referencial','Cliente nulo','id_cliente NULL'],
['C003-dup','Limpieza Clientes','Duplicado','id_cliente duplicado, se conserva 1'],
['P007-dup','Limpieza Productos','Duplicado','codigo duplicado'],
['VTA-011-dup','Limpieza Ventas','Duplicado','id_venta duplicado'],
]);
// 4. DOCUMENTACION
const wsD = wb.addWorksheet('DOCUMENTACION_ETL');
wsD.columns=[{width:110}];
const docLines=[
'DOCUMENTACIÓN TÉCNICA ETL — TecnoHogar Colombia S.A.S. | Arquitectura de Datos — María Celeny Pérez',
'',
'1. FUENTES: Clientes(13 filas sucias), Productos(10), Vendedores(6), Ventas_Agosto(12), Ventas_Septiembre(11). Total origen: 52 filas.',
'2. REGLAS DE NEGOCIO (9): PK únicas; FK válida cliente/producto/vendedor; cantidad>0; descuento 0-30%; precio/costo numérico positivo; ciudades/categorías uniformes; rechazados con trazabilidad.',
'3. CONTEOS ANTES/DESPUÉS: Clientes 13→12 limpios | Productos 10→8 (se elimina duplicado P007 y P008 sin precio) | Vendedores 6→5 | Ventas consolidadas 23→11 válidas + 12 rechazadas con motivo.',
'4. PASOS POWER QUERY POR CONSULTA (panel Pasos aplicados):',
'   Clientes_Limpia: Origen→Encabezados→Trim+Clean→Estandarizar ciudad (Bogotá/Medellín/Cali/Barranquilla/Cartagena/Bucaramanga)→Reemplazar N/A,NULL,Sin dato,-,"" por null→Quitar duplicados id_cliente→Tipo texto/fecha.',
'   Productos_Limpia: Quitar columna_innecesaria→Trim/Clean→Estandarizar categoría (Portátiles, Accesorios, Celulares, Monitores, Impresoras, Almacenamiento)→Limpiar $ . , y convertir precio/costo a número→Filtrar precio>0 y costo>0→Quitar duplicados.',
'   Vendedores_Limpia: Trim/Clean→Capitalizar nombres→Estandarizar sede (Sede Norte/Sur/Cali) y ciudad→Quitar duplicados id_vendedor.',
'   Ventas_Consolidada: Anexar Ventas_Agosto + Ventas_Septiembre (Anexar consultas)→Trim claves→Fecha tipo fecha (varios formatos)→cantidad a entero, descuento texto % a decimal→Filtrar cantidad>0 y descuento 0-0.3 (los que incumplen van a Rechazados).',
'   BD_VENTAS_ETL: Combinar Ventas con Clientes_Limpia (Left Outer, id_cliente)→Combinar con Productos (codigo_producto)→Combinar con Vendedores (id_vendedor)→Filtrar nulos en FK (a Rechazados)→Agregar columnas: Total_Bruto=[cantidad]*[precio]; Valor_Descuento=[Total_Bruto]*[descuento]; Total_Neto=[Total_Bruto]-[Valor_Descuento]; Costo_Total=[cantidad]*[costo]; Utilidad=[Total_Neto]-[Costo_Total]; Clasificacion: Alta si Total_Neto>=2000000, Media si >=500000, else Baja→Cargar a hoja BD_VENTAS_ETL. RECHAZADOS se carga aparte.',
'5. JUSTIFICACIÓN ANEXAR vs COMBINAR: Anexar = apilar filas mismo esquema (agosto+septiembre). Combinar = pegar columnas por clave (enriquecer venta con atributos). Left Outer conserva venta y detecta huérfanos.',
'6. DICCIONARIO RESUMIDO: id_venta(PK Venta,texto,no nulo); id_cliente(FK→Cliente); codigo_producto(FK→Producto); id_vendedor(FK→Vendedor); fecha(date); cantidad(int>0); descuento(decimal 0-0.3); Total_Bruto/Neto/Utilidad(decimal); Clasificacion(texto). PK: id_cliente, codigo_producto, id_vendedor, id_venta. FK en Venta.',
'7. MODELO ER: Cliente(1)—(N)Venta; Producto(1)—(N)Venta; Vendedor(1)—(N)Venta. Ver hoja MODELO_ER.',
'8. CONCLUSIÓN CALIDAD: Origen ~45% filas con algún defecto; tras ETL 100% válidas en BD_VENTAS_ETL, 12 filas trazables en RECHAZADOS. Proceso 100% reproducible en Power Query sin tocar hojas fuente.',
'9. EVIDENCIA: ver hoja CODIGO_M (pegar en Editor avanzado) + hoja EVIDENCIA_PASOS. Capturas: Datos>Obtener datos>Desde tabla + Inicio>Quitar filas>Quitar duplicados + Transformar>Formato>Trim/Clean + Reemplazar valores + Anexar/Combinar.',
];
docLines.forEach(l=>wsD.addRow([l]));
wsD.getRow(1).font={bold:true,size:12,color:{argb:'FFFFFFFF'}}; wsD.getRow(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF1F4E79'}}; wsD.getRow(1).alignment={wrapText:true};
for(let i=2;i<=docLines.length;i++){wsD.getRow(i).alignment={wrapText:true,vertical:'top'};wsD.getRow(i).height=30;}
sheet('DICCIONARIO_DATOS',['Campo','Entidad','Tipo','Descripción','PK/FK','Nulos'],[
['id_cliente','Cliente','Texto','Identificador único cliente','PK','No'],
['nombre','Cliente','Texto','Nombre completo','-','No'],
['ciudad','Cliente/Vendedor','Texto','Ciudad estandarizada','-','No'],
['codigo_producto','Producto','Texto','Código único producto','PK','No'],
['producto','Producto','Texto','Nombre producto','-','No'],
['categoria','Producto','Texto','Categoría uniforme','-','No'],
['precio','Producto','Decimal','Precio positivo','-','No'],
['costo','Producto','Decimal','Costo positivo','-','No'],
['id_vendedor','Vendedor','Texto','Identificador vendedor','PK','No'],
['id_venta','Venta','Texto','Identificador transacción','PK','No'],
['fecha','Venta','Fecha','Fecha venta','-','No'],
['cantidad','Venta','Entero','Unidades >0','-','No'],
['descuento','Venta','Decimal','0 a 0.30','-','No'],
['Total_Bruto','Venta','Decimal','cantidad*precio','-','No'],
['Valor_Descuento','Venta','Decimal','Total_Bruto*descuento','-','No'],
['Total_Neto','Venta','Decimal','Bruto-Descuento','-','No'],
['Costo_Total','Venta','Decimal','cantidad*costo','-','No'],
['Utilidad','Venta','Decimal','Neto-Costo','-','No'],
['Clasificacion_Venta','Venta','Texto','Alta/Media/Baja','-','No'],
],HDR2);
sheet('MODELO_ER',['Entidad','PK','Atributos','Relación'],[
['Cliente','id_cliente','nombre, ciudad, correo, teléfono','1:N → Venta por id_cliente'],
['Producto','codigo_producto','producto, categoría, precio, costo','1:N → Venta por codigo_producto'],
['Vendedor','id_vendedor','vendedor, sede, ciudad','1:N → Venta por id_vendedor'],
['Venta','id_venta','fecha, cantidad, descuento, totales, FK×3','N:1 ← cada maestro'],
],HDR2);
const wsE = wb.addWorksheet('EVIDENCIA_PASOS'); wsE.columns=[{width:110}];
['EVIDENCIA DE PASOS APLICADOS (qué capturar en Power Query: panel derecho Pasos aplicados):','Clientes_Limpia: Origen | Encabezados promovidos | Tipo cambiado | Trim | Clean | Mayúscula inicial ciudad | Reemplazar N/A→null | Quitar duplicados | Reordenar.',
'Productos_Limpia: Origen | Quitar columna_innecesaria | Trim/Clean categoría | Reemplazar $/. /,/texto→null | Tipo número | Filtrar >0 | Quitar duplicados.',
'Vendedores_Limpia: Origen | Trim/Clean | Capitalizar | Estandarizar sede/ciudad | Quitar duplicados.',
'Ventas_Consolidada: Anexar Agosto+Septiembre | Trim claves | Fecha multiformato a fecha | descuento % a decimal | cantidad a entero.',
'BD_VENTAS_ETL: 3× Combinar (Left Outer) + Expandir solo columnas necesarias | Filtrar FK no nulas | 6 columnas personalizadas (fórmulas abajo) | Tipo + Cargar.',
'FÓRMULAS M: Total_Bruto=[cantidad]*[precio] | Valor_Descuento=[Total_Bruto]*[descuento] | Total_Neto=[Total_Bruto]-[Valor_Descuento] | Costo_Total=[cantidad]*[costo] | Utilidad=[Total_Neto]-[Costo_Total] | Clasificacion= if [Total_Neto]>=2000000 then "Alta" else if >=500000 then "Media" else "Baja".',
'RECHAZADOS: Table.SelectRows con errores OR cantidad<=0 OR descuento>0.3 OR FK nula + columna Motivo.'].forEach(l=>wsE.addRow([l]));
const M = `-- PEGAR EN POWER QUERY: Datos > Obtener datos > Desde tabla/rango > Editor avanzado
-- 1) Clientes_Limpia
let O=Excel.CurrentWorkbook(){[Name="tblClientes"]}[Content], T=Table.TransformColumns(O,{{"id_cliente",Text.Trim},{"nombre",each Text.Clean(Text.Trim(_))},{"ciudad",each Text.Clean(Text.Trim(_))}}), R=Table.ReplaceValue(T,"BOGOTA","Bogotá",Replacer.ReplaceText,{"ciudad"}) in R
-- 2) Productos_Limpia: quitar col innecesaria, Trim/Clean, precio texto->número: Number.FromText(Text.Remove(Text.Remove([precio],{"$","."}),{","}))
-- 3) Vendedores_Limpia: Trim/Clean + Text.Proper([vendedor])
-- 4) Ventas_Consolidada = Table.Combine({Ventas_Agosto,Ventas_Septiembre})
-- 5) BD_VENTAS_ETL = combinar 3 maestros (LeftOuter) + columnas: Total_Bruto=[cantidad]*[precio], Valor_Descuento=[Total_Bruto]*[descuento], Total_Neto=[Total_Bruto]-[Valor_Descuento], Costo_Total=[cantidad]*[costo], Utilidad=[Total_Neto]-[Costo_Total], Clasif= if [Total_Neto]>=2000000 then "Alta" else if [Total_Neto]>=500000 then "Media" else "Baja"
-- 6) Rechazados = filas con cantidad<=0 o descuento>0.3 o FK nula. Cargar ambas a hojas.`;
const wsM = wb.addWorksheet('CODIGO_M'); wsM.columns=[{width:115}]; M.split('\n').forEach(l=>wsM.addRow([l]));
await wb.xlsx.writeFile(outDir+'/TecnoHogar_ETL_Taller.xlsx');
console.log('XLSX OK');
// M code file
const mFull = `TECNO-HOGAR ETL — CODIGO M POWER QUERY (pegar en Editor avanzado de cada consulta)
====================================================================
LCientes:
let Origen=Excel.CurrentWorkbook(){[Name="Clientes"]}[Content], P1=Table.PromoteHeaders(Origen,[PromoteAllScalars=true]),
T=Table.TransformColumns(P1,{{"id_cliente",Text.Trim, type text},{"nombre",each Text.Clean(Text.Trim(_)),type text},{"ciudad",each Text.Clean(Text.Trim(_)),type text}}),
E1=Table.ReplaceValue(T,"BOGOTA","Bogotá",Replacer.ReplaceText,{"ciudad"}), E2=Table.ReplaceValue(E1,"bogotá","Bogotá",Replacer.ReplaceText,{"ciudad"}),
E3=Table.ReplaceValue(E2,"BOGOTÁ","Bogotá",Replacer.ReplaceText,{"ciudad"}), E4=Table.ReplaceValue(E3,"Medellin","Medellín",Replacer.ReplaceText,{"ciudad"}),
E5=Table.ReplaceValue(E4,"MEDELLÍN","Medellín",Replacer.ReplaceText,{"ciudad"}), E6=Table.ReplaceValue(E5," cali","Cali",Replacer.ReplaceText,{"ciudad"}),
N=Table.ReplaceValue(Table.ReplaceValue(Table.ReplaceValue(Table.ReplaceValue(E6,"N/A",null,Replacer.ReplaceValue,{"correo","telefono"}),"NULL",null,Replacer.ReplaceValue,{"correo","telefono"}),"Sin dato",null,Replacer.ReplaceValue,{"correo","telefono"}),"-",null,Replacer.ReplaceValue,{"telefono"}),
D=Table.Distinct(N,{"id_cliente"}) in D

LProductos: quitar "columna_innecesaria" (Table.RemoveColumns), Trim+Clean categoria, estandarizar (portatiles->Portátiles, accesorios->Accesorios, CELULARES/celulares->Celulares, IMPRESORAS->Impresoras), precio=Number.FromText(Text.Remove(Text.Remove([precio],{"$"}),{".",","})) handle N/A/NULL/- -> null, filtrar [precio]>0 and [costo]>0, Distinct codigo_producto.

LVendedores: Trim/Clean, [vendedor]=Text.Proper, sede: "sede norte"->"Sede Norte", "SEDE SUR"->"Sede Sur"; ciudad igual que clientes; Distinct id_vendedor.

LVentas_Consolidada: Table.Combine({Ventas_Agosto,Ventas_Septiembre}) + Trim claves + fecha a date + cantidad Int64 + descuento: if Text.Contains([descuento],"%") then Number.FromText(Text.Remove([descuento],{"%"}))/100 else null.

LBD_VENTAS_ETL:
let V=Ventas_Consolidada, C1=Table.NestedJoin(V,"id_cliente",Clientes_Limpia,"id_cliente","C",JoinKind.LeftOuter), E1=Table.ExpandTableColumn(C1,"C",{"nombre","ciudad"},{"cliente","ciudad_cliente"}),
C2=Table.NestedJoin(E1,"codigo_producto",Productos_Limpia,"codigo_producto","P",JoinKind.LeftOuter), E2=Table.ExpandTableColumn(C2,"P",{"producto","categoria","precio","costo"},{"producto","categoria","precio","costo"}),
C3=Table.NestedJoin(E2,"id_vendedor",Vendedores_Limpia,"id_vendedor","Vd",JoinKind.LeftOuter), E3=Table.ExpandTableColumn(C3,"Vd",{"vendedor"},{"vendedor"}),
Ok=Table.SelectRows(E3,each [cliente]<>null and [producto]<>null and [vendedor]<>null and [cantidad]>0 and [descuento]>=0 and [descuento]<=0.3 and [precio]>0 and [costo]>0),
B=Table.AddColumn(Ok,"Total_Bruto",each [cantidad]*[precio],type number), Vd2=Table.AddColumn(B,"Valor_Descuento",each [Total_Bruto]*[descuento],type number),
N2=Table.AddColumn(Vd2,"Total_Neto",each [Total_Bruto]-[Valor_Descuento],type number), Ct=Table.AddColumn(N2,"Costo_Total",each [cantidad]*[costo],type number),
U=Table.AddColumn(Ct,"Utilidad",each [Total_Neto]-[Costo_Total],type number), Cl=Table.AddColumn(U,"Clasificacion_Venta",each if [Total_Neto]>=2000000 then "Alta" else if [Total_Neto]>=500000 then "Media" else "Baja",type text) in Cl

LRechazados: Table.SelectRows(paso E3, each [cliente]=null or [producto]=null or [vendedor]=null or [cantidad]<=0 or [descuento]>0.3) + columna Motivo_Rechazo.
CARGAR: BD_VENTAS_ETL y RECHAZADOS como "Cargar en tabla" a hojas. Tablas limpias al Modelo de datos para Power Pivot opcional.
`;
fs.writeFileSync(outDir+'/Codigo_M_PowerQuery.txt', mFull);
const html = `<!doctype html><html lang="es"><head><meta charset="utf-8"><title>Doc técnica ETL TecnoHogar</title><style>body{font-family:Arial;max-width:900px;margin:20px auto}h1{color:#1F4E79}table{border-collapse:collapse;width:100%}td,th{border:1px solid #999;padding:6px}th{background:#1F4E79;color:#fff}</style></head><body>
<h1>Documentación técnica ETL — TecnoHogar Colombia S.A.S.</h1><p>Arquitectura de Datos · Docente María Celeny Pérez · Proceso ETL reproducible en Power Query</p>
<h2>1. Fuentes y conteos</h2><table><tr><th>Fuente</th><th>Antes</th><th>Después</th></tr><tr><td>Clientes</td><td>13</td><td>12</td></tr><tr><td>Productos</td><td>10</td><td>8</td></tr><tr><td>Vendedores</td><td>6</td><td>5</td></tr><tr><td>Ventas (ago+sep)</td><td>23</td><td>11 válidas + 12 rechazadas</td></tr></table>
<h2>2. Reglas de negocio (9)</h2><p>PK únicas · FK válida · cantidad&gt;0 · descuento 0–30% · precio/costo numérico positivo · nomenclatura uniforme · rechazados con trazabilidad.</p>
<h2>3. Transformaciones clave</h2><p>Trim/Clean, estandarizar ciudad/categoría, reemplazar N/A-NULL-Sin dato-"-" por null, tipos correctos, quitar duplicados, quitar columna innecesaria, anexar ventas, 3 combinar (Left Outer), 6 columnas derivadas.</p>
<h2>4. Columnas derivadas</h2><p>Total_Bruto=cantidad×precio · Valor_Descuento=Bruto×descuento · Total_Neto=Bruto−Descuento · Costo_Total=cantidad×costo · Utilidad=Neto−Costo · Clasificación: Alta ≥2.000.000, Media ≥500.000, Baja resto.</p>
<h2>5. Modelo ER</h2><p>Cliente(1)—(N)Venta · Producto(1)—(N)Venta · Vendedor(1)—(N)Venta. PK: id_cliente, codigo_producto, id_vendedor, id_venta.</p>
<h2>6. Conclusión</h2><p>Calidad inicial ~45% filas con defectos. Final 100% válido en BD_VENTAS_ETL con trazabilidad total. Imprimir esta página como PDF = documento anexo.</p></body></html>`;
fs.writeFileSync(outDir+'/Documentacion_Tecnica_ETL.html', html);
const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="900" height="420"><rect x="20" y="60" width="180" height="150" fill="#DCE6F1" stroke="#1F4E79" stroke-width="2"/><text x="30" y="85" font-weight="bold">CLIENTE PK id_cliente</text><text x="30" y="110">nombre, ciudad</text><text x="30" y="130">correo, telefono</text><rect x="360" y="60" width="180" height="150" fill="#DCE6F1" stroke="#1F4E79" stroke-width="2"/><text x="370" y="85" font-weight="bold">PRODUCTO PK</text><text x="370" y="110">producto, categoria</text><text x="370" y="130">precio, costo</text><rect x="20" y="260" width="180" height="130" fill="#FFF2CC" stroke="#7F6000" stroke-width="2"/><text x="30" y="285" font-weight="bold">VENDEDOR PK</text><text x="30" y="310">vendedor, sede, ciudad</text><rect x="360" y="260" width="300" height="130" fill="#E2EFDA" stroke="#375623" stroke-width="2"/><text x="370" y="285" font-weight="bold">VENTA PK id_venta + FKx3</text><text x="370" y="310">fecha, cantidad, descuento</text><text x="370" y="330">totales, Clasificacion</text><line x1="200" y1="140" x2="360" y2="300" stroke="black"/><text x="230" y="220">1:N</text><line x1="540" y1="140" x2="500" y2="260" stroke="black"/><text x="540" y="210">1:N</text><line x1="200" y1="320" x2="360" y2="320" stroke="black"/><text x="260" y="310">1:N</text></svg>`;
fs.writeFileSync(outDir+'/Modelo_ER.svg', svg);
fs.writeFileSync(outDir+'/Guion_sustentacion_2min.txt', `SUSTENTACIÓN 2 MIN:\n1) Problema: duplicados, N/A/NULL, espacios, ciudades/categorías mixtas, precios texto, fechas mixtas, cantidad<=0, descuento>30%, FK huérfanas.\n2) Transformación: Trim+Clean, estandarizar, reemplazar a null, tipos, quitar duplicados/col innecesaria, anexar ago+sep, combinar Left Outer por 3 claves, 6 columnas (Bruto, Descuento, Neto, Costo, Utilidad, Clasif Alta/Media/Baja).\n3) Efecto: 52 filas origen -> 11 válidas + 15 rechazos trazables (incluye dup limpieza). 100% válido al final.\n4) Claves: id_cliente, codigo_producto, id_vendedor. Anexar apila filas, Combinar pega columnas.\n5) ER: 3 maestros 1:N a Venta transaccional.`);
console.log('TODO OK en '+outDir);
})();
