# -*- coding: utf-8 -*-
import os

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = os.path.join(CURRENT_DIR, "imagenes")
if not os.path.exists(IMG_DIR):
    IMG_DIR = r'C:\Users\fsilv\OneDrive\Documentos\GitHub\Novela\imagenes'

def _img(name):
    return os.path.join(IMG_DIR, name)

ILUSTRACIONES_POR_CAPITULO = {
    3: {'illustration': _img('kaoruko_waguri.png'), 'illustration_caption': ''},
    4: {'illustration': _img('subaru_hoshina.png'), 'illustration_caption': 'Ilustración: Subaru Hoshina (Clase F-2)'},
    5: {'illustration': _img('ken.png'), 'illustration_caption': 'Ilustración: Ken (Clase F-2)'},
    8: {'illustration': _img('hijos_de_calydon.png'), 'illustration_caption': 'Ilustración: Los Hijos de Calydon (Clase F-2)'},
    23: {'illustration': _img('ken_fenix.png'), 'illustration_caption': 'Ilustración: Ken desatando el Espíritu del Fénix (Clase F-2)'},
    37: {'illustration': _img('lortelle_kecheln_herida.png'), 'illustration_caption': 'Ilustración: Lortelle Kecheln (Caída en el Abismo)'},
    41: {'illustrations': [{'image': _img('vlad.png'), 'caption': 'Ilustración: Vlad (Caballero Aprendiz del Ducado)'}, {'image': _img('agnes_tachyon.png'), 'caption': 'Ilustración: Agnes Tachyon (Maga Experimental de la Academia)'}]},
    44: {'illustrations': [{'image': _img('akame.png'), 'caption': 'Ilustración: Akame (La Hoja Silenciosa de la F-2)'}, {'image': _img('rukia_renji.jpg'), 'caption': 'Ilustración: Rukia y Renji (Dúo de Hielo y Furia)'}]},
    45: {'illustrations': [{'image': _img('subaru_natsuki.png'), 'caption': 'Ilustración: Natsuki Subaru (Despertar del Aura de Batalla)'}]},
    46: {'illustrations': [{'image': _img('hijos_de_calydon.png'), 'caption': 'Ilustración: Los Hijos de Calydon (Operación de Rescate)'}]},
    47: {'illustrations': [{'image': _img('erza_scarlet.png'), 'caption': 'Ilustración: Erza Scarlet (Defensa Implacable del Flanco)'}, {'image': _img('akame.png'), 'caption': 'Ilustración: Akame (El Filo Silencioso de la Victoria)'}]},
    48: {'illustrations': [{'image': _img('kaoruko_waguri.png'), 'caption': 'Ilustración: Kaoruko Waguri (La Calidez Tras la Tormenta)'}, {'image': _img('lortelle_kecheln.png'), 'caption': 'Ilustración: Lortelle Kecheln (Recuperada y Lista para el Verano)'}]},
    49: {'illustrations': [{'image': _img('zemina.png'), 'caption': 'Ilustración: Zemina (La Invitada Especial de Vlad)'}]},
    50: {'illustrations': [{'image': _img('vlad.png'), 'caption': 'Ilustración: Vlad (Capitán Táctico de la Clase B)'}, {'image': _img('zemina.png'), 'caption': 'Ilustración: Zemina (Amor y Destino en la Costa Esmeralda)'}]},
    51: {'illustrations': [{'image': _img('kaoruko_waguri.png'), 'caption': 'Ilustración: Kaoruko Waguri (Dulzura y Calidez en la Noche Marina)'}, {'image': _img('rukia_renji.jpg'), 'caption': 'Ilustración: Renji y Rukia (Confidencias Bajo la Luz Lunar)'}]},
    52: {'illustrations': [{'image': _img('subaru_natsuki.png'), 'caption': 'Ilustración: Natsuki (El Espíritu Pragmático y Empático de la F-2)'}, {'image': _img('hijos_de_calydon.png'), 'caption': 'Ilustración: Los Hijos de Calydon (Lucy y Piper Wheel)'}]},
    53: {'illustrations': [{'image': _img('lortelle_kecheln.png'), 'caption': 'Ilustración: Lortelle Kecheln (Seducción y Pasión en la Costa Esmeralda)'}]},
    54: {'illustrations': [{'image': _img('lortelle_kecheln.png'), 'caption': 'Ilustración: Lortelle Kecheln (Complicidad al Amanecer)'}, {'image': _img('kaoruko_waguri.png'), 'caption': 'Ilustración: Kaoruko Waguri (El Desayuno de la Clase F-2)'}]},
    55: {'illustrations': [{'image': _img('kaoruko_waguri.png'), 'caption': 'Ilustración: Kaoruko Waguri (La Dulzura de la Entrega)'}]},
    56: {'illustrations': [{'image': _img('elysea_de_vane.png'), 'caption': 'Ilustración: Lady Elysea de Vane (Bajo el Brillo de la Medianoche)'}]},
    57: {'illustrations': [{'image': _img('natsuki_lucy_playa.png'), 'caption': 'Ilustración: Natsuki y Lucy (Complicidad, Sandía y Sonrojos en la Orilla)'}]},
    59: {'illustrations': [{'image': _img('costa_esmeralda_vacaciones.jpg'), 'caption': 'Ilustración: Seier, Kaoruko, Lortelle y Elysea en la Terraza de Costa Esmeralda'}]},
    60: {'illustrations': [{'image': _img('natsuki_besa_lucy.jpg'), 'caption': 'Ilustración: Natsuki y Lucy (Un Beso Dulce y Sincero Bajo el Ocaso)'}]},
    81: {'illustrations': [{'image': _img('lortelle_kecheln.png'), 'caption': 'Ilustración: Lortelle Kecheln (La Seda Dorada y la Inteligencia de Elte)'}]},
    82: {'illustrations': [{'image': _img('erza_scarlet.png'), 'caption': 'Ilustración: Erza Scarlet (Baluarte Acorazado y Disciplina en la División Beta)'}]},
    83: {'illustrations': [{'image': _img('kaoruko_waguri.png'), 'caption': 'Ilustración: Kaoruko Waguri (La Bendición Sagrada Antes de la Marcha al Sur)'}]},
    84: {'illustrations': [{'image': _img('elysea_de_vane.png'), 'caption': 'Ilustración: Lady Elysea de Vane (El Broche de Escarcha y la Despedida Noble)'}, {'image': _img('ken_fenix.png'), 'caption': 'Ilustración: Ken (El Quinto Eslabón y el Fuego Solar)'}]},
    85: {'illustrations': [{'image': _img('agnes_tachyon.png'), 'caption': 'Ilustración: Agnes Tachyon (Inspección Alquímica y Prueba de Respiradores en el Frente Sur)'}]},
    86: {'illustrations': [{'image': _img('erza_scarlet.png'), 'caption': 'Ilustración: Erza Scarlet (Despliegue de la Emperatriz del Fuego en el Abismo)'}, {'image': _img('ken.png'), 'caption': 'Ilustración: Ken (El Golpe Solar del Fénix y la Purificación del Miasma)'}]},
}


