# -*- coding: utf-8 -*-
import os

NOVELA_DIR = r'C:\Users\fsilv\OneDrive\Documentos\GitHub\Novela'

ILUSTRACIONES_POR_CAPITULO = {
    3: {'illustration': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\kaoruko_waguri.png', 'illustration_caption': ''},
    4: {'illustration': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\subaru_hoshina.png', 'illustration_caption': 'Ilustración: Subaru Hoshina (Clase F-2)'},
    5: {'illustration': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\ken.png', 'illustration_caption': 'Ilustración: Ken (Clase F-2)'},
    8: {'illustration': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\hijos_de_calydon.png', 'illustration_caption': 'Ilustración: Los Hijos de Calydon (Clase F-2)'},
    23: {'illustration': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\ken_fenix.png', 'illustration_caption': 'Ilustración: Ken desatando el Espíritu del Fénix (Clase F-2)'},
    37: {'illustration': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\lortelle_kecheln_herida.png', 'illustration_caption': 'Ilustración: Lortelle Kecheln (Caída en el Abismo)'},
    41: {'illustrations': [{'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\vlad.png', 'caption': 'Ilustración: Vlad (Caballero Aprendiz del Ducado)'}, {'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\agnes_tachyon.png', 'caption': 'Ilustración: Agnes Tachyon (Maga Experimental de la Academia)'}]},
    44: {'illustrations': [{'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\akame.png', 'caption': 'Ilustración: Akame (La Hoja Silenciosa de la F-2)'}, {'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\rukia_renji.jpg', 'caption': 'Ilustración: Rukia y Renji (Dúo de Hielo y Furia)'}]},
    45: {'illustrations': [{'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\subaru_natsuki.png', 'caption': 'Ilustración: Natsuki Subaru (Despertar del Aura de Batalla)'}]},
    46: {'illustrations': [{'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\hijos_de_calydon.png', 'caption': 'Ilustración: Los Hijos de Calydon (Operación de Rescate)'}]},
    47: {'illustrations': [{'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\erza_scarlet.png', 'caption': 'Ilustración: Erza Scarlet (Defensa Implacable del Flanco)'}, {'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\akame.png', 'caption': 'Ilustración: Akame (El Filo Silencioso de la Victoria)'}]},
    48: {'illustrations': [{'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\kaoruko_waguri.png', 'caption': 'Ilustración: Kaoruko Waguri (La Calidez Tras la Tormenta)'}, {'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\lortelle_kecheln.png', 'caption': 'Ilustración: Lortelle Kecheln (Recuperada y Lista para el Verano)'}]},
    49: {'illustrations': [{'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\zemina.png', 'caption': 'Ilustración: Zemina (La Invitada Especial de Vlad)'}]},
    50: {'illustrations': [{'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\vlad.png', 'caption': 'Ilustración: Vlad (Capitán Táctico de la Clase B)'}, {'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\zemina.png', 'caption': 'Ilustración: Zemina (Amor y Destino en la Costa Esmeralda)'}]},
    51: {'illustrations': [{'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\kaoruko_waguri.png', 'caption': 'Ilustración: Kaoruko Waguri (Dulzura y Calidez en la Noche Marina)'}, {'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\rukia_renji.jpg', 'caption': 'Ilustración: Renji y Rukia (Confidencias Bajo la Luz Lunar)'}]},
    52: {'illustrations': [{'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\subaru_natsuki.png', 'caption': 'Ilustración: Natsuki (El Espíritu Pragmático y Empático de la F-2)'}, {'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\hijos_de_calydon.png', 'caption': 'Ilustración: Los Hijos de Calydon (Lucy y Piper Wheel)'}]},
    53: {'illustrations': [{'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\lortelle_kecheln.png', 'caption': 'Ilustración: Lortelle Kecheln (Seducción y Pasión en la Costa Esmeralda)'}]},
    54: {'illustrations': [{'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\lortelle_kecheln.png', 'caption': 'Ilustración: Lortelle Kecheln (Complicidad al Amanecer)'}, {'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\kaoruko_waguri.png', 'caption': 'Ilustración: Kaoruko Waguri (El Desayuno de la Clase F-2)'}]},
    55: {'illustrations': [{'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\kaoruko_waguri.png', 'caption': 'Ilustración: Kaoruko Waguri (La Dulzura de la Entrega)'}]},
    56: {'illustrations': [{'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\elysea_de_vane.png', 'caption': 'Ilustración: Lady Elysea de Vane (Bajo el Brillo de la Medianoche)'}]},
    57: {'illustrations': [{'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\natsuki_lucy_playa.png', 'caption': 'Ilustración: Natsuki y Lucy (Complicidad, Sandía y Sonrojos en la Orilla)'}]},
    59: {'illustrations': [{'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\costa_esmeralda_vacaciones.jpg', 'caption': 'Ilustración: Seier, Kaoruko, Lortelle y Elysea en la Terraza de Costa Esmeralda'}]},
    60: {'illustrations': [{'image': 'C:\\Users\\fsilv\\OneDrive\\Documentos\\GitHub\\Novela\\imagenes\\natsuki_besa_lucy.jpg', 'caption': 'Ilustración: Natsuki y Lucy (Un Beso Dulce y Sincero Bajo el Ocaso)'}]},
}
