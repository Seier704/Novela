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
    "num": 78,
    "header": "Capítulo 78",
    "title": "Crónicas del Despliegue: Ecos de Fuego, Sombras y Acero",
    "paragraphs": [
        "El fuego del hogar de *La Jarra Plateada* crepitaba con fuerza, proyectando sombras danzarinas sobre las vigas de roble y las jarras de estaño.",
        "A medida que la noche avanzaba y los relatos de Morvath dejaban paso al tintineo de los cubiertos, apoyé los codos sobre la mesa de la cabecera. Miré a cada uno de mis amigos con una media sonrisa retadora, cruzando miradas con Ken, Caesar, Subaru, Kaoruko y Erza:",
        "—Ya les contamos con lujo de detalles cómo Vlad, Akame y la Sección Uno le arrancamos las fauces al culto en el sur... —dije dando un trago a mi cerveza—. Pero no crean que se librarán de su parte. ¿Qué hicieron ustedes durante este mes mientras nosotros nos arrastrábamos por alcantarillas y fundiciones?",
        "Ken fue el primero en reaccionar. El artista marcial golpeó la mesa con su jarra, haciendo temblar los platos entre carcajadas:",
        "—¡Ja! ¡Nosotros no nos quedamos calentando asientos, hermano! ¡Nuestra misión en la Garganta de Fuego fue puro infierno!",
        "* * *",
        "Ken se acomodó la bandana roja en la frente y comenzó su relato con los ojos chispeantes de euforia marcial.",
        "Había sido asignado a la vanguardia de las estribaciones volcánicas del este junto a un contingente mixto de las Clases B y C, bajo la supervisión remota de su maestra Hwa Young. Al principio, los cadetes aristócratas miraban con desdén a Ken por pelear a puño limpio sin escudo ni peto de acero, tratándolo como un campesino insensato.",
        "Todo cambió cuando un Drake de Magma Alfa —un reptil colosal de escamas de basalto y aliento ígneo de dos mil grados— emboscó a una caravana de refugiados mineros en un desfiladero angosto.",
        "—Los nobles con armaduras pesadas casi se cocinan vivos dentro de sus corazas por la radiación térmica —narraba Ken gesticulando con entusiasmo—. La bestia iba a calcinar las carretas con los niños. No tuve tiempo de esperar órdenes: activé el paso del Fénix, salté sobre las rocas incandescentes y me metí directo bajo la mandíbula del lagarto.",
        "Ken relató cómo desvió un coletazo con la palma y concentró toda su energía solar en un gancho ascendente demoledor: el impacto reventó las placas óseas de la bestia en una detonación de llamas doradas que la derribó al fondo del barranco. Al final de la expedición, los mismos nobles que se burlaban de él terminaron cargándole las provisiones en señal de respeto absoluto.",
        "—¡Un golpe limpio y seco! —celebró Borgar chocando su jarra con la de Ken—. ¡A los dragones se les parte la quijada de abajo hacia arriba, muchacho! ¡Bien hecho!",
        "* * *",
        "A mi lado, Kaoruko sonrió con timidez, acariciando el dorso de mi mano antes de tomar la palabra animada por Yami.",
        "—A nosotras nos tocó una misión muy distinta, pero igual de peligrosa —explicó Kaoruko con su voz dulce y templada—. Lady Vane y yo fuimos asignadas a la protección logística del Valle de los Lirios, custodiando un convoy de semillas, boticarios y trigo para los pueblos del este.",
        "En mitad de los bosques templados, una banda de cuarenta desertores mercenarios intentó asaltar las carretas para apoderarse de los suministros medicinales.",
        "—Eran hombres violentos y desesperados —continuó Kaoruko—. Pero no queríamos provocar una carnicería. Cuando cargaron a caballo, levanté una barrera sagrada entrelazada con raíces de tierra que desarmó a los jinetes al atrapar las cinchas sin lastimar a los animales. Luego, Lady Vane avanzó con su estoque de escarcha:",
        "La dama aristocrática, con su porte gélido y letal, congeló los aceros y las empuñaduras de los bandidos al mero contacto, inutilizando sus armas antes de que pudieran derramar una sola gota de sangre. Capturaron a la banda completa y entregaron el cargamento intacto, recibiendo coronas de flores y bendiciones de los ancianos del valle.",
        "—Control perfecto del maná —asintió Vlad con admiración—. Neutralizar sin masacrar requiere diez veces más destreza que aniquilar.",
        "* * *",
        "Caesar King soltó una carcajada estruendosa, mostrando una mella reciente en el borde de su gran escudo blindado:",
        "—¡Nosotros no tuvimos tanta delicadeza, te lo aseguro, Vlad! ¡El Paso del Trueno estaba infestado de Mantícoras y Guivernos rabiosos!",
        "Yami, recostada con elegancia sobre el respaldo de su silla, tomó el relevo con una sonrisa socarrona:",
        "—Caesar hacía de cebo viviente. Se plantaba en mitad del peñasco, rugía más fuerte que los monstruos y aguantaba las embestidas aéreas con su escudo de asedio como si fuera una muralla de hierro. En cuanto la mantícora quedaba clavada en sus defensas, yo caía desde los riscos a sus espaldas.",
        "Yami narró cómo deslizaba sus dagas dobles imbuidas en miasma sombrío, cortando las alas membranosas en pleno vuelo y decapitando a las bestias antes de que pudieran clavar sus aguijones venenosos. La cacería fue tan aplastante que concluyeron el contrato tres días antes de lo previsto, bajando la carne de los monstruos al pueblo más cercano para armar un asado comunal con cerveza de montaña.",
        "—Pura adrenalina de trinchera —comentó Vance con los ojos brillantes, maravillado por la sincronía entre el escudo de Caesar y las sombras de Yami.",
        "* * *",
        "Natsuki Subaru se puso de pie de un salto, apoyando un pie en su banco con su clásica pose heroica y dramática:",
        "—¡Y no olviden la gesta de los Pantanos de Niebla-Negra! ¡Una odisea de lodo, veneno y gloria táctica!",
        "Lucy, sentada a su lado, resopló divertida tapándose el rostro con una mano:",
        "—No exageres tanto, Subaru... aunque admito que fue un desastre al principio.",
        "Subaru explicó cómo su grupo fue enviado a rastrear ciénagas brumosas infestadas de alimañas Skavens y acechadores del fango, donde una patrulla de exploradores de tercer año había quedado atrapada en arenas movedizas. Mientras Lucy utilizaba sus dagas gemelas para eliminar con destreza letal a los hostigadores en la niebla, Piper Wheel modificó el motor a vapor de una carreta cisterna, colocándole orugas flotantes de madera.",
        "—¡Subaru usó bombas de fósforo y silbatos señuelo para confundir a los monstruos en los juncos! —reveló Lucy mirando a su compañero con orgullo—. Nos abrió una ruta limpia y logramos sacar a los seis cadetes heridos antes de que se hundieran en el pantano.",
        "—El cerebro también es un arma, amigos míos —bromeó Subaru hinchando el pecho, provocando risas y aplausos de toda la mesa.",
        "* * *",
        "Erza Scarlet fue la última en compartir su informe, haciéndolo con la sobriedad marcial que la caracterizaba.",
        "Había sido desplegada en las Colinas Plateadas junto a Agnes Tachyon para investigar una perturbación arcana en una cripta subterránea del Imperio Antiguo. El lugar estaba custodiado por autómatas de granito y mithril reactivados por una sobrecarga de maná puro.",
        "—Los constructos eran resistentes al daño cortante convencional —explicó Erza con mirada serena—. Tuve que recurrir a la Armadura de la Rueda del Cielo para contener a diez golems a la vez con espadas flotantes en círculos defensivos, mientras Agnes analizaba los circuitos rúnicos.",
        "Agnes se acomodó los anteojos con suficiencia intelectual:",
        "—Eran matrices de control de la Tercera Dinastía. Fascinantes. Logré revertir la polaridad de los cristales de alimentación en menos de veinte minutos, neutralizando los núcleos sin destruir las reliquias históricas. Un éxito académico total.",
        "* * *",
        "Cuando el último relato terminó, el salón quedó envuelto en un silencio cálido y reverente.",
        "Miré a cada uno de los rostros congregados alrededor de las jarras vacías. Ken con sus nudillos humeantes; Kaoruko y Lucy con sus sonrisas llenas de ternura y templanza; Yami y Caesar con su lealtad forjada en fuego; Vlad y Vance representando la verdadera nobleza de espada; Subaru, Erza, Agnes y nuestros nuevos hermanos de Morvath.",
        "Ninguno de nosotros era el mismo novato inexperto que había cruzado las puertas de la Academia meses atrás.",
        "Teníamos diecisiete años recién cumplidos. Habíamos derramado sangre, salvado campesinos, derrotado herejías y desafiado a la muerte en los cuatro puntos cardinales del reino. La vieja discriminación entre nobles y plebeyos, entre la Clase A y la F, se había pulverizado por completo bajo el peso de nuestros actos.",
        "Alcé mi jarra rebosante hacia el centro de la mesa. Todos siguieron el gesto de inmediato, haciendo chocar el metal con un estruendo formidable que retumbó en la taberna:",
        "—La academia quería probarnos enviándonos a los confines del mapa —dije con voz firme y profunda—. Y lo único que lograron fue hacernos más fuertes.",
        "—¡POR LA GENERACIÓN DORADA! —rugió Ken alzando su puño envuelto en chispas.",
        "Los brindis y las risas se prolongaron hasta la madrugada bajo el cielo estrellado de Solaria. La misión de largo alcance había terminado, pero la leyenda de nuestra hermandad apenas comenzaba a escribirse."
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
    guardar_capitulo_individual(CAPITULO, NOVELA_DIR)
    compilar_libro_completo(NOVELA_DIR)
    print("===============================================================")
    print(" ¡PROCESO COMPLETADO EXITOSAMENTE!")
    print("===============================================================")

if __name__ == "__main__":
    ejecutar()
