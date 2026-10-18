# -*- coding: utf-8 -*-
import os
import sys

NOVELA_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Novela"
PROYECTO_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Proyecto-Minecraft\Novela"

sys.path.insert(0, NOVELA_DIR)
if os.path.exists(PROYECTO_DIR):
    sys.path.insert(0, PROYECTO_DIR)

from generar_libro import create_novel_document, create_single_chapter_document
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

CAPITULO_45 = {
    "header": "Capítulo 45",
    "title": "Puños de Fuego y Voluntad Plateada",
    "illustrations": [
        {
            "image": os.path.join(NOVELA_DIR, "imagenes", "subaru_natsuki.png"),
            "caption": "Ilustración: Natsuki Subaru (Despertar del Aura de Batalla)"
        }
    ],
    "paragraphs": [
        "Con la retaguardia asegurada por el filo silencioso de Akame y el santuario gélido de Rukia, Kaoruko y yo corrimos de regreso hacia el frente sin concedernos un solo respiro.",
        "El portón principal de la fortaleza era un auténtico hervidero de metal, chispas y sangre. La caída de las tres primeras Ratas Ogro no había detenido a la horda; al contrario, al ver que la brecha exterior seguía abierta, los señores de la guerra Skaven habían ordenado el avance de sus batallones más sanguinarios: las **Alimañas de la Tormenta**, cubiertas de armaduras pesadas de placas remachadas, y equipos de apoyo equipados con pesados **Lanzallamas de Disformidad**.",
        "Un chorro colosal de fuego verde y humeante barrió la pasarela de la barricada, haciendo arder los sacos de arena y forzando a los lanceros cadetes a retroceder tosiendo por el calor corrosivo.",
        "—¡Mantengan la distancia! ¡Ese fuego consume el metal! —gritó Vlad desde su posición de mando, desviando una andanada de flechas con su espada bastarda mientras intentaba reorganizar a la infantería.",
        "Dos operarios Skaven, con tanques de gas presurizado atados a sus jorobas, se adelantaron entre las ruinas del rastrillo, riendo con chillidos histéricos mientras preparaban un segundo disparo concentrado hacia los cadetes arrinconados.",
        "No alcanzaron a abrir la válvula.",
        "* * *",
        "—¡Oigan, pedazos de basura pestilente! ¡En esta fortaleza el único que tiene permitido prender fuego soy yo! —una voz arrogante, vibrante y llena de una confianza contagiosa retumbó desde lo alto del arco de piedra.",
        "Una llamarada carmesí cruzó el aire nocturno como un cometa.",
        "Lighter cayó en picada en mitad de la brecha.",
        "Aterrizó en cuclillas con una precisión felina, levantando una nube de chispas ardientes. Llevaba su chaqueta de combate desabrochada, su pañuelo ondeando con insolencia y sus nudillos reforzados con guanteletes mecánicos que silbaban con vapor sobrecalentado. Su sonrisa descarada brillaba bajo la luz de las llamas.",
        "El primer operador del lanzallamas intentó girar la boquilla hacia él, pero Lighter se movió más rápido que el parpadeo de un ojo.",
        "Con el balanceo y desvío de un boxeador de élite, esquivó el cañón metálico por el exterior, penetró en la guardia de la criatura y descargó un gancho de izquierda demoledor directo a su hígado. El impacto de su puño, envuelto en fuego condensado, pulverizó la coraza de hierro del Skaven como si fuera cáscara de huevo.",
        "—¡Uno! —exclamó Lighter con una sonrisa salvaje.",
        "Sin frenar el giro de sus caderas, encadenó un volado de derecha cargado de combustión directa que impactó contra el tanque de disformidad a la espalda del segundo operario.",
        "*¡¡KRA-BOOOOOOM!!*",
        "El contenedor explotó en una llamarada cegadora, incinerando a una docena de alimañas que venían detrás y abriendo un claro humeante en la entrada.",
        "Cinco Alimañas de la Tormenta, armadas con alabardas de dos metros, intentaron ensartar a Lighter simultáneamente en el cuello de botella. Pero el púgil de la Clase F-2 estaba en su salsa: agachándose con un péndulo perfecto, resbaló entre las astas de las armas y comenzó a soltar una ráfaga implacable de rectos, ganchos y uppercuts que destrozaban mandíbulas acorazadas al ritmo de explosiones cortas de maná ígneo.",
        "—¡Vamos, ratitas! ¿Es todo lo que tienen para el campeón? —bramaba Lighter, haciendo girar sus hombros con una soltura formidable mientras dejaba un sendero de cuerpos humeantes a su paso.",
        "* * *",
        "Sin embargo, el flanco derecho de la barricada corría peligro de ser desbordado.",
        "Ocho guerreros Skaven de élite, aprovechando que la atención de Lighter estaba en el centro, treparon por un montón de escombros para caer sobre la retaguardia de los tiradores.",
        "Fue en ese instante crítico donde una presencia brillante y decidida interceptó el asalto.",
        "Natsuki dio tres zancadas explosivas hacia el montículo de ruinas.",
        "Natsuki ya no luchaba con la torpeza del recién llegado ni con la desesperación ciega del superviviente. Durante las últimas semanas de entrenamiento espartano en la F-2 y bajo la tutela brutal de Yami y Seier, su cuerpo había cambiado, pero lo que realmente había madurado era su mente.",
        "Cerró los ojos por una milésima de segundo, regulando su respiración: *Inhala... exhala... siente el peso del aire, anticipa la sombra antes de que corte.*",
        "Su **Aura de Batalla** reaccionó a su estado mental.",
        "La energía plateada con destellos índigo que envolvía su figura no solo creció en densidad; se volvió fluida, moldeándose a sus músculos como una segunda piel etérea. Un zumbido agudo resonó en el aire cuando la percepción de Natsuki se expandió de golpe, sincronizando sus reflejos con la trayectoria de cada arma enemiga.",
        "La primera Alimaña de la Tormenta descargó un tajo descendente con su hacha pesada. Natsuki no intentó bloquear la fuerza bruta: dio un paso lateral milimétrico, permitiendo que la hoja se clavara en la tierra a centímetros de su bota, y con una fluidez pasmosa clavó su daga corta envuelta en aura plateada directo en la juntura del cuello del monstruo.",
        "El impacto de su maná condensado no solo atravesó la carne, sino que detonó en una onda concéntrica que proyectó al enemigo hacia atrás.",
        "Dos guerreros más se abalanzaron sobre él en pinza.",
        "—¡No me subestimen! —rugió Natsuki.",
        "Canalizó su aura hacia la palma de su mano libre y la proyectó hacia adelante en una ráfaga de choque improvisada, desestabilizando el centro de gravedad del asaltante de la izquierda, mientras con un giro gimnástico esquivaba la lanza del de la derecha y le fracturaba la rodilla de una patada reforzada.",
        "Cada movimiento de Natsuki era más eficiente, más letal y más seguro que el anterior. La duda y el miedo habían desaparecido por completo de sus ojos, reemplazados por el instinto templado de un guerrero que comprendía el valor de su propia vida y la de sus camaradas.",
        "El último asaltante intentó abalanzarse con una daga traicionera; Natsuki lo interceptó con un barrido implacable, enviándolo por los aires directo hacia el radio de acción de Lighter.",
        "—¡Envío especial, Lighter! —gritó Natsuki con la respiración agitada pero con una sonrisa firme.",
        "—¡Recibido con gusto, novato! —respondió Lighter, conectando un uppercut ascendente que mandó a la rata a volar contra las vigas superiores del portón.",
        "Lighter y Natsuki cruzaron miradas y chocaron los puños en mitad de la polvareda.",
        "—Vaya, vaya... parece que te estás convirtiendo en un verdadero monstruo del combate cuerpo a cuerpo, Natsuki —comentó Lighter con una sonrisa de aprobación franca.",
        "—Tengo buenos maestros... y no pienso quedarme atrás de ninguno de ustedes —respondió Natsuki, con su aura plateada aún brillando con un fulgor orgulloso y constante.",
        "* * *",
        "Aprovechando la brecha abierta por Lighter y Natsuki, Kaoruko y yo nos estrellamos contra el centro del contingente enemigo.",
        "Descargué una combinación fulminante de patada baja, rodillazo al plexo y codazo circular reforzado con maná de tierra, destrozando las líneas Skaven como un ariete humano. A mi derecha, la hoja de Kaoruko danzaba como un rayo plateado, abriendo caminos limpios con una elegancia letal que dejaba a los atacantes sin tiempo de levantar sus defensas.",
        "Con Vlad cerrando los flancos mediante la falange de escudos de los cadetes, Yami cortando las oleadas con tajos de gravedad oscura y Ken proyectando ráfagas de fuego sagrado desde el torreón este, el ímpetu de la horda finalmente se quebró.",
        "Cientos de ratas comenzaron a arrojar sus armas y a retroceder en pánico hacia el desfiladero, tropezando con los cadáveres de sus propios hermanos.",
        "Los vítores de los cadetes comenzaron a resonar en las murallas.",
        "Sin embargo, antes de que pudiéramos cantar victoria, un cuerno de guerra titánico, forjado con el cráneo de un leviatán de piedra bruja, rugió desde la niebla profunda del cañón con una vibración tan pesada que hizo crujir las piedras de los cimientos.",
        "Las ratas que huían se detuvieron en seco, arrojándose al fango temblando de terror.",
        "Desde la oscuridad del desfiladero, cabalgando sobre una monstruosidad colosal acorazada con púas de hierro y colmillos de dos metros, emergió el verdadero comandante del nido: un gigantesco **Señor de la Guerra Skaven**, cuya alabarda rúnica ardía con una llama verde que amenazaba con devorar la noche entera.",
        "La batalla había llegado a su punto culminante.",
        "Y el líder del abismo venía a cobrar cada vida con su propia mano."
    ]
}

def compilar_capitulo_45():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 45              ")
    print("==================================================")
    
    from compilar_novela import PROLOGO, CAPITULO_1
    from agregar_capitulo_2 import CAPITULO_2
    from agregar_capitulo_3 import CAPITULO_3
    from agregar_capitulo_4 import CAPITULO_4
    from agregar_capitulo_5 import CAPITULO_5
    from agregar_capitulo_6 import CAPITULO_6
    from agregar_capitulo_7 import CAPITULO_7
    from agregar_capitulo_8 import CAPITULO_8
    from agregar_capitulo_9 import CAPITULO_9
    from agregar_capitulo_10 import CAPITULO_10
    from agregar_capitulo_11 import CAPITULO_11
    from agregar_capitulo_12 import CAPITULO_12
    from agregar_capitulo_13 import CAPITULO_13
    from agregar_capitulo_14 import CAPITULO_14
    from agregar_capitulo_15 import CAPITULO_15
    from agregar_capitulo_16 import CAPITULO_16
    from agregar_capitulo_17 import CAPITULO_17
    from agregar_capitulo_18 import CAPITULO_18
    from agregar_capitulo_19 import CAPITULO_19
    from agregar_capitulo_20 import CAPITULO_20
    from agregar_capitulo_21 import CAPITULO_21
    from agregar_capitulo_22 import CAPITULO_22
    from agregar_capitulo_23 import CAPITULO_23
    from agregar_capitulo_24 import CAPITULO_24
    from agregar_capitulo_25 import CAPITULO_25
    from agregar_capitulo_26 import CAPITULO_26
    from agregar_capitulo_27 import CAPITULO_27
    from agregar_capitulo_28 import CAPITULO_28
    from agregar_capitulo_29 import CAPITULO_29
    from agregar_capitulo_30 import CAPITULO_30
    from agregar_capitulo_31 import CAPITULO_31
    from agregar_capitulo_32 import CAPITULO_32
    from agregar_capitulo_33 import CAPITULO_33
    from agregar_capitulo_34 import CAPITULO_34
    from agregar_capitulo_35 import CAPITULO_35
    from agregar_capitulo_36 import CAPITULO_36
    from agregar_capitulo_37 import CAPITULO_37
    from agregar_capitulo_38 import CAPITULO_38
    from agregar_capitulo_39 import CAPITULO_39
    from agregar_capitulo_40 import CAPITULO_40
    from agregar_capitulo_41 import CAPITULO_41
    from agregar_capitulo_42 import CAPITULO_42
    from agregar_capitulo_43 import CAPITULO_43
    from agregar_capitulo_44 import CAPITULO_44
    
    titulo = "El camino del reencarnado"
    subtitulo = "Crónicas de Sangre, Acero y Deseo"
    autor = "Seier Silva"
    capitulos = [
        PROLOGO, CAPITULO_1, CAPITULO_2, CAPITULO_3, CAPITULO_4,
        CAPITULO_5, CAPITULO_6, CAPITULO_7, CAPITULO_8, CAPITULO_9,
        CAPITULO_10, CAPITULO_11, CAPITULO_12, CAPITULO_13, CAPITULO_14,
        CAPITULO_15, CAPITULO_16, CAPITULO_17, CAPITULO_18, CAPITULO_19,
        CAPITULO_20, CAPITULO_21, CAPITULO_22, CAPITULO_23, CAPITULO_24,
        CAPITULO_25, CAPITULO_26, CAPITULO_27, CAPITULO_28, CAPITULO_29,
        CAPITULO_30, CAPITULO_31, CAPITULO_32, CAPITULO_33, CAPITULO_34,
        CAPITULO_35, CAPITULO_36, CAPITULO_37, CAPITULO_38, CAPITULO_39,
        CAPITULO_40, CAPITULO_41, CAPITULO_42, CAPITULO_43, CAPITULO_44,
        CAPITULO_45
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 45
    txt_cap45 = os.path.join(capitulos_dir, "45_Capitulo_45.txt")
    with open(txt_cap45, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_45['header'].upper()}: {CAPITULO_45['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_45['paragraphs']))
    print(f"[OK] Texto plano Capítulo 45 guardado en:\n  {txt_cap45}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 45
    docx_cap45 = os.path.join(capitulos_dir, "45_Capitulo_45.docx")
    create_single_chapter_document(CAPITULO_45, docx_cap45, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 45 guardado en:\n  {docx_cap45}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 45)
    master_docx = os.path.join(NOVELA_DIR, "compilado", "El_Camino_del_Reencarnado.docx")
    try:
        create_novel_document(titulo, subtitulo, autor, capitulos, master_docx)
        print(f"[OK] Libro Maestro actualizado en:\n  {master_docx}")
    except PermissionError:
        print(f"[AVISO] '{master_docx}' está abierto en Microsoft Word.")
        backup_master = os.path.join(NOVELA_DIR, "compilado", "El_Camino_del_Reencarnado_temp.docx")
        create_novel_document(titulo, subtitulo, autor, capitulos, backup_master)
        print(f"[OK] Guardado en copia auxiliar:\n  {backup_master}")

if __name__ == "__main__":
    compilar_capitulo_45()
