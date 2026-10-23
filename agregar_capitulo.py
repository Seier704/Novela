# -*- coding: utf-8 -*-
"""
===============================================================================
GESTOR UNIFICADO DE CAPÍTULOS - "EL CAMINO DEL REENCARNADO"
===============================================================================
Este archivo es la plantilla GENERAL y ÚNICA para agregar o sobrescribir capítulos.
Para agregar un nuevo capítulo, simplemente edite el diccionario `CAPITULO` con
el número, título y párrafos correspondientes y ejecute este script:
    python agregar_capitulo.py

El script automáticamente:
1. Genera el documento Word individual: capitulos/XX_Capitulo_XX.docx
2. Recompila el libro maestro completo: compilado/El_Camino_del_Reencarnado.docx
===============================================================================
"""

import os
import sys
import glob
import re
import docx

# Definición de rutas base
NOVELA_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Novela"
PROYECTO_DIR = r"c:\Users\fsilv\OneDrive\Documentos\GitHub\Proyecto-Minecraft\Novela"

for d in [NOVELA_DIR, PROYECTO_DIR]:
    if os.path.exists(d) and d not in sys.path:
        sys.path.insert(0, d)

from generar_libro import create_novel_document, create_single_chapter_document

# ===============================================================================
# DEFINICIÓN DEL CAPÍTULO A AGREGAR / ACTUALIZAR
# ===============================================================================
CAPITULO = {
    "num": 81,
    "header": "Capítulo 81",
    "title": "El Engranaje de Astraea: Tres Generaciones y la Seda Dorada",
    "paragraphs": [
        "Tras la conmoción de la asamblea inaugural en el Gran Anfiteatro Rúnico, el mediodía se abrió paso sobre la Academia de Astraea con un sol templado que hacía resplandecer los techos de pizarra azul y las fuentes de mármol.",
        "Los pasillos, claustros y explanadas se convirtieron en un hervidero de pasos apresurados y conversaciones animadas. El nuevo Sistema de Formación Integral había sacudido los cimientos de la institución, pero bastaba recorrer la Gran Calzada de los Cedros Blancos para comprender que la academia no era un aula estática, sino un ecosistema vibrante donde tres generaciones de cadetes convivían bajo un mismo cielo.",
        "En los campos de maniobra inferiores, el bullicio pertenecía por entero a los recién llegados de **Primer Año**.",
        "Eran chicos y chicas de quince y dieciséis años, con uniformes recién desempacados que todavía olían a tinte y lana limpia. Con la abolición de las castas, la marea de novatos era un mosaico fascinante: hijos de labradores que sostenían sus primeras espadas de madera con dedos nudosos, jóvenes magas de cuna plebeya que intentaban canalizar chispas elementales en varitas de fresno —provocando cómicas explosiones de humo negro que las hacían toser entre risas nerviosas—, y tímidos aprendices semihumanos y enanos que miraban con asombro las monumentales torres de la capital.",
        "Para todos ellos, los cadetes de Segundo Año que cruzábamos los jardines con charreteras plateadas éramos poco menos que figuras de leyenda.",
        "Cada vez que Vlad, Akame o yo pasábamos cerca de las barandillas, los novatos se congelaban a mitad de sus prácticas. Los cuchicheos y las miradas de reverencia se extendían como un reguero de pólvora: sabían perfectamente que éramos los supervivientes del asedio fronterizo y los héroes que habían purgado el feudo de Morvath. Para aquellos muchachos, que durante años solo habían escuchado sobre la tiranía de los nobles, ver a plebeyos y soldados condecorados al mando de la academia era un faro de esperanza.",
        "En mitad del camino, un cadete novato de complexión delgada que cargaba un cajón pesado de piedras rúnicas para el laboratorio de transmutación tropezó con una raíz de cedro. El chico cayó de bruces y los cristales rodaron por el empedrado.",
        "—¡A-Ay, no, no, por favor...! —balbuceó el muchacho, palideciendo de terror mientras recogía las piedras con manos temblorosas, esperando instintivamente un insulto o una patada como solían hacer los antiguos cadetes de la A-1.",
        "Me detuve y me agaché a su lado. Con mi mano enguantada recogí tres de los cristales más pesados, los deposité con suavidad en el cajón y luego le tendí la diestra al novato, ayudándolo a incorporarse de un solo tirón firme y templado:",
        "—El equilibrio en las piernas se gana con el entrenamiento, recluta —le dije con una sonrisa serena y voz profunda—. Camina con el peso distribuido en el centro de gravedad y no bajes la mirada. En esta academia nadie te castigará por tropezar; solo te exigiremos que te levantes con más fuerza.",
        "El muchacho se quedó sin habla, con los ojos húmedos de asombro y gratitud pura, antes de cuadrarse de golpe y llevarse el puño al pecho con un saludo militar perfecto:",
        "—¡S-Sí, señor Silva! ¡Muchas gracias, mi señor!",
        "Vlad, que caminaba a mi lado con paso marcial, sonrió de medio lado con aprobación:",
        "—Hace un año nos habrían escupido por caminar por este sendero, Seier. Es reconfortante ver que la semilla que plantamos en Morvath y en la frontera empieza a florecer.",
        "* * *",
        "Pero la academia no solo miraba hacia los novatos. Al cruzar hacia la Gran Calzada de los Héroes, el aire adquirió una densidad muy distinta.",
        "Allí se desplazaban los cadetes de **Tercer Año**: los últimos de la escala estudiantil, jóvenes de dieciocho y diecinueve años que estaban a escasas semanas de sus pruebas finales de graduación para recibir despachos oficiales como tenientes de las legiones imperiales, inquisidores de la corte o comandantes de bastión.",
        "Vestían uniformes oscuros con dobles charreteras doradas, capas ribeteadas de terciopelo y armas rúnicas de forja maestra ceñidas al cinto. Desprendían un aura pesada, curtida y solemne. Eran los veteranos que habían visto pasar promociones enteras y que se preparaban para marchar a la verdadera guerra en las fronteras exteriores del imperio.",
        "Un escuadrón de tercer año se detuvo frente a nosotros al pie de la escalinata central. Su líder era un espadachín de casi metro noventa, de rostro severo y cicatriz en la mandíbula: el Capitán Bawden, uno de los mejores duelistas de la promoción saliente.",
        "El silencio que se formó entre ambos grupos fue tenso, pero no hostil. Bawden recorrió con la mirada mi corpulencia, la calma letal de Akame con la mano descansando sobre *Murasame* y la medalla del León de Oro en mi pecho. Tras unos segundos de evaluación marcial, el veterano de tercer año alzó su guantelete y golpeó su coraza en un saludo de profundo respeto entre pares:",
        "—Silva. El informe confidencial sobre la caída del culto en Morvath llegó a manos del Estado Mayor la semana pasada —dijo Bawden con voz ronca y respetuosa—. Hicieron un trabajo formidable en las fundiciones del este. Muchos en mi promoción creían que los cadetes de segundo año eran demasiado blandos para la guerra real, pero ustedes demostraron tener más agallas que muchos caballeros con título.",
        "—Solo cumplimos con nuestro deber, capitán —respondí devolviendo el saludo con firmeza militar.",
        "—Nuestra graduación es inminente —añadió Bawden con una media sonrisa antes de reanudar el paso—. Cuando nosotros marchemos a los puestos fronterizos del norte este invierno, Astraea quedará bajo su custodia. Prepárense bien para las misiones de alto riesgo que Gerald les asignará. El reino va a necesitar oficiales de acero, no niños con pergaminos.",
        "Los vi marcharse con sus capas ondeando bajo el viento. La jerarquía de la academia era clara: los de primer año aprendían a gatear, los de segundo año éramos el martillo táctico en pleno ascenso, y los de tercer año eran los guardianes que estaban a punto de saltar al abismo del mundo exterior.",
        "* * *",
        "—Vaya, vaya... Miren al orgullo de Solaria. Un paseo de diez minutos por el campus y ya tienes a los novatos rezándote como a un santo y a los de tercer año reconociéndote como su heredero militar.",
        "Una voz sedosa, melodiosa y cargada de una picardía inconfundible brotó desde la sombra de los arcos de jazmín del ala este.",
        "Giré la cabeza y una sonrisa involuntaria curvó mis labios al encontrarme con **Lortelle Kecheln**.",
        "Estaba apoyada con indolente elegancia contra una balaustrada de mármol, sosteniendo un abanico de plumas esmeralda. Su cabello castaño claro con reflejos dorados caía en cascadas sedosas sobre sus hombros, enmarcando un rostro aristocrático de belleza felina y unos ojos verdes centelleantes que desbordaban inteligencia y coquetería. Llevaba un conjunto entallado de terciopelo verde bosque con un chaleco de brocado bordado en hilo de oro de la Compañía Elte, resaltando con descarada sofisticación las curvas de su busto y su cintura esbelta.",
        "—Vlad, Akame... si me disculpan, tengo asuntos mercantiles urgentes que tratar con su capitán —dijo Lortelle dedicándoles un guiño pícaro.",
        "Vlad sonrió con disimulo y Akame simplemente asintió mientras masticaba una tira de carne seca. Me acerqué a ella, pero antes de que pudiera pronunciar palabra, Lortelle me tomó suavemente de la muñeca y tiró de mí hacia el interior del pabellón consular de la Compañía Elte dentro de la academia.",
        "Cruzamos una antecámara lujosa decorada con tapices persas y alfombras espesas hasta entrar en su despacho privado. En cuanto la pesada puerta de roble se cerró, Lortelle activó con un toque de su anillo un glifo rúnico de aislamiento acústico total.",
        "Apenas la barrera mágica parpadeó en azul, la máscara de diplomática altiva de Lortelle se evaporó por completo.",
        "Giró sobre sus talones, rodeó mi cuello con ambos brazos y se alzó sobre las puntas de sus pies para estampar sus labios contra los míos. Fue un beso voraz, cálido y profundo, impregnado de ese perfume a sándalo y canela que me transportó de inmediato a las noches desinhibidas que habíamos compartido en la Costa Esmeralda. Sentí su cuerpo suave y ardiente presionarse contra el mío, mientras sus dedos acariciaban mi nuca con familiaridad posesiva.",
        "Cuando finalmente se separó, sus mejillas estaban teñidas de un tenue rubor y su respiración era deliciosamente entrecortada:",
        "—No sabes cuánto te extrañé durante esas malditas semanas en Morvath, mi bruto favorito... —susurró rozando la comisura de mis labios con el pulgar—. Aunque me alegra ver que volviste en una sola pieza.",
        "Lortelle se acomodó en el borde de su gran escritorio de caoba pulida, cruzando una pierna sobre la otra con sensualidad calculada mientras me servía una copa de licor de pera helado:",
        "—Pero no solo te traje aquí para robarte besos, Seier. Hay negocios colosales en marcha y tú eres el corazón de todos ellos.",
        "—Sabía que la Compañía Elte no tardaría en tejer su red en este segundo año —bromeé dando un trago a la copa.",
        "Lortelle sonrió con esa mirada felina que la hacía tan peligrosa como fascinante:",
        "—El nuevo sistema interdisciplinario de Gerald es una mina de oro. He estructurado la economía de Astraea en tres frentes perfectos:",
        "Lortelle comenzó a enumerar con precisión magistral:",
        "—Primero: los de **Primer Año**. La Compañía Elte ha financiado cincuenta becas completas para los plebeyos y forasteros más talentosos. Les entregamos kits de iniciación, reactivos alquímicos y armaduras de cuero a bajo costo. Para ellos soy una benefactora divina; para la compañía, significa que en dos años tendremos cincuenta oficiales de élite con contratos exclusivos de lealtad a Elte.",
        "—Segundo: los de **Tercer Año**. Los egresados están gastando auténticas fortunas en pertrechos de graduación. Les estoy vendiendo lingotes de oricalco, matrices de mithril y catalizadores de alta pureza al por mayor. Y a cambio de mis descuentos, sus secretarios me entregan informes confidenciales sobre las rutas de patrulla y las vacantes en los regimientos imperiales.",
        "—Y tercero... lo más importante: **ustedes, los de Segundo Año**.",
        "Lortelle extrajo de un cajón secreto un pergamino sellado con cera azul de la Cancillería de la Corona y me lo extendió con expresión solemne:",
        "—Mis contactos en el Estado Mayor me pasaron esto antes de que Gerald lo anuncie. Los **Escuadrones de Despliegue de Alto Riesgo** no van a ser enviados a simples patrullas rurales. La Corona detectó fluctuaciones masivas de miasma en las Ruinas Negras del Sur y una reactivación de nidos Skavens en las viejas fallas subterráneas. Están reservando a los mejores de segundo año para aplastar esas amenazas antes de que se conviertan en una invasión.",
        "Leí el informe con detenimiento. Los datos tácticos eran oro puro: rutas de acceso, densidad de maná contaminado y patrones de movimiento enemigo.",
        "Alcé la mirada hacia Lortelle. La joven heredera se inclinó hacia mí, apoyando ambas manos sobre mis hombros y mirándome a los ojos con devoción sincera y complicidad absoluta:",
        "—Tú pones el acero, la fuerza y el liderazgo en el campo de batalla, Seier... pero yo me encargaré de que nunca te falte oro, suministros ni la mejor información del reino. Nadie tocará a mi hombre ni a su escuadrón mientras yo controle los hilos tras bambalinas.",
        "La tomé por la cintura con firmeza, atrayéndola hacia mi regazo para darle un beso tierno y prolongado que la hizo suspirar de satisfacción.",
        "—Eres increíble, Lortelle —le dije mirándola con sincero afecto—. No podría pedir una mejor aliada ni una mujer más brillante a mi lado.",
        "—Y más te vale no olvidarlo —ronroneó ella acariciando mi pecho con una sonrisa pícara—. Ahora prepárate, capitán. Con los novatos aprendiendo, los veteranos despidiéndose y nosotros al frente, este segundo año va a ser nuestra mayor conquista."
    ]
}

# ===============================================================================
# FUNCIONES DE GESTIÓN Y COMPILACIÓN
# ===============================================================================

def guardar_capitulo_individual(cap_dict, base_dir=NOVELA_DIR):
    """Guarda el archivo .docx individual del capítulo."""
    capitulos_dir = os.path.join(base_dir, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    num = cap_dict.get("num", 0)
    if num == 0 or "PROLOGO" in cap_dict.get("header", "").upper() or "PRÓLOGO" in cap_dict.get("header", "").upper():
        base_name = "00_Prologo"
    else:
        base_name = f"{num:02d}_Capitulo_{num}"
        
    docx_path = os.path.join(capitulos_dir, f"{base_name}.docx")
    create_single_chapter_document(
        cap_dict,
        docx_path,
        novel_title="El camino del reencarnado",
        author="Seier Silva"
    )
    print(f"[OK] Documento Word individual guardado:\n  {docx_path}")
    return docx_path


def cargar_todos_los_capitulos(base_dir=NOVELA_DIR):
    """Lee todos los capítulos directamente desde la carpeta capitulos/*.docx."""
    capitulos_dir = os.path.join(base_dir, "capitulos")
    docx_files = sorted(glob.glob(os.path.join(capitulos_dir, "*.docx")))
    
    # Intentar cargar mapa de ilustraciones
    ilustraciones = {}
    try:
        from ilustraciones_capitulos import ILUSTRACIONES_POR_CAPITULO
        ilustraciones = ILUSTRACIONES_POR_CAPITULO
    except Exception:
        pass
        
    lista_capitulos = []
    
    for fpath in docx_files:
        filename = os.path.basename(fpath)
        num = None
        m_num = re.match(r"^(\d+)_", filename)
        if m_num:
            num = int(m_num.group(1))
            
        doc = docx.Document(fpath)
        if not doc.paragraphs:
            continue
            
        title_p = doc.paragraphs[0].text.strip()
        parts = title_p.split("\n")
        header = parts[0].strip() if len(parts) > 0 else "Capítulo"
        title = parts[1].strip() if len(parts) > 1 else ""
        
        paragraphs = []
        for p in doc.paragraphs[1:]:
            t = p.text.strip()
            if not t:
                continue
            if t in ["— ✦ —", "— ◆ —", "✦ ✦ ✦", "— ✧ —"]:
                continue
            if t.startswith("Ilustración:") or t.startswith("Ilustracion:"):
                continue
            paragraphs.append(t)
            
        chap_dict = {
            "header": header,
            "title": title,
            "paragraphs": paragraphs
        }
        
        # Asociar ilustraciones
        if num in ilustraciones:
            illus_info = ilustraciones[num]
            if "illustrations" in illus_info:
                chap_dict["illustrations"] = illus_info["illustrations"]
            elif "illustration" in illus_info:
                chap_dict["illustration"] = illus_info["illustration"]
                chap_dict["illustration_caption"] = illus_info.get("illustration_caption", "")
                
        lista_capitulos.append(chap_dict)
        
    return lista_capitulos


def compilar_libro_completo(base_dir=NOVELA_DIR):
    """Recompila el libro maestro completo a partir de todos los .docx de capitulos/."""
    titulo = "El camino del reencarnado"
    subtitulo = "Crónicas de Sangre, Acero y Deseo"
    autor = "Seier Silva"
    
    capitulos = cargar_todos_los_capitulos(base_dir)
    print(f"[*] Compilando Libro Maestro con {len(capitulos)} capítulos (desde .docx)...")
    
    compilado_dir = os.path.join(base_dir, "compilado")
    os.makedirs(compilado_dir, exist_ok=True)
    master_docx = os.path.join(compilado_dir, "El_Camino_del_Reencarnado.docx")
    
    try:
        create_novel_document(titulo, subtitulo, autor, capitulos, master_docx)
        print(f"[OK] Libro Maestro actualizado exitosamente en:\n  {master_docx}")
    except PermissionError:
        print(f"[AVISO] '{master_docx}' está abierto en Microsoft Word.")
        backup_master = os.path.join(compilado_dir, "El_Camino_del_Reencarnado_temp.docx")
        create_novel_document(titulo, subtitulo, autor, capitulos, backup_master)
        print(f"[OK] Guardado en copia auxiliar:\n  {backup_master}")


def ejecutar():
    print("===============================================================")
    print(f" AGREGAR / ACTUALIZAR: {CAPITULO['header']} - {CAPITULO['title']}")
    print("===============================================================")
    for base in [NOVELA_DIR, PROYECTO_DIR]:
        if os.path.exists(base):
            print(f"[*] Procesando en: {base}")
            guardar_capitulo_individual(CAPITULO, base)
            compilar_libro_completo(base)
    print("===============================================================")
    print(" ¡PROCESO COMPLETADO EXITOSAMENTE!")
    print("===============================================================")

if __name__ == "__main__":
    ejecutar()
