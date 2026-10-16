# -*- coding: utf-8 -*-
import os
from generar_libro import create_novel_document, create_single_chapter_document


PROLOGO = {
    "header": "Prólogo",
    "title": "Cenizas y Juramento",
    "paragraphs": [
        "El último sonido que escuché de mi vida pasada fue el crujido desgarrador del metal cediendo a diez mil metros de altura.",
        "Un fallo en las turbinas. La cabina despresurizándose con un aullido ensordecedor. Las luces de emergencia parpadeando en un rojo frenético mientras las máscaras de oxígeno caían del techo. Recuerdo el agarre tembloroso de mi madre a mi izquierda, los nudillos blancos de mi padre intentando calmar a mi hermana menor, y la sensación de vacío en el estómago cuando el avión se precipitó en una caída libre ingobernable hacia el océano.",
        "No hubo héroes. No hubo aterrizajes milagrosos ni supervivientes. Solo un impacto brutal que pulverizó el fuselaje, seguido de una negrura absoluta.",
        "Y luego... el despertar.",
        "* * *",
        "Abrí los ojos en una habitación de piedra tosca, iluminada por la luz vacilante de un candil de sebo. Mis manos no eran las mías: eran diminutas, pálidas, pertenecientes a un niño de apenas seis años.",
        "Tardé semanas en procesar lo que había ocurrido. No era un sueño febril ni una alucinación del más allá. Había reencarnado conservando cada uno de mis recuerdos intactos, en un mundo de corte medieval, áspero y despiadado, donde el maná fluía por el aire y las bestias acechaban más allá de las murallas.",
        "En esta nueva existencia, mi nombre era Seier Silva. Único hijo y heredero de la Casa Baronal Silva.",
        "Mis padres en este mundo, Ronald y Elena Silva, no eran los típicos aristócratas decadentes que se pavoneaban en banquetes perfumados. Eran veteranos con cicatrices en el rostro, guerreros de frontera que gobernaban un bastión rocoso en los confines del sur del reino. Su deber era simple y brutal: ser el primer escudo de la humanidad contra las incursiones que descendían de las Tierras Sombrías.",
        "A pesar de la aspereza del clima y la austeridad de la fortaleza, en su hogar encontré un calor familiar genuino. Mi padre me sentaba en sus rodillas y me mostraba los mapas tácticos de los desfiladeros; mi madre me envolvía en pieles gruesas mientras cantaba canciones antiguas frente a la chimenea.",
        "Pero la paz en la frontera dura tan poco como una vela bajo la tormenta.",
        "* * *",
        "Ocurrió un año después de mi séptimo invierno.",
        "Las campanas de alarma sonaron en mitad de la noche con un tañido fúnebre. Las atalayas arremetían con fuegos de bengala carmesí. No era una simple incursión de trasgos o lobos de sombra; era una horda demoníaca organizada, bestias acorazadas con fauces chorreantes de miasma negro que rompieron las empalizadas como ramas secas.",
        "Mi padre había enviado jinetes de emergencia días antes a las tierras vecinas del Marqués Delacroix solicitando las tropas de apoyo estipuladas por decreto real. Los jinetes jamás regresaron.",
        "Años más tarde comprendería la amarga verdad: los nobles de la corte, envidiosos del prestigio militar de los Silva y codiciosos de nuestras canteras de hierro, retrasaron deliberadamente los refuerzos. Querían que la Casa Silva fuera borrada del mapa.",
        "—¡Abran las compuertas traseras! ¡Evacúen a los artesanos, a las mujeres y a los niños por el paso de montaña! —bramó Ronald Silva con la cota de malla bañada en sangre negra.",
        "—¡Mi señor, no resistirán el puente ustedes solos! —suplicó el capitán de la guardia.",
        "—¡No estoy pidiendo opiniones, soldado! ¡Cumpla la orden!",
        "Recuerdo a mi madre, con los ojos vidriosos pero la mandíbula firme, empujándome hacia el interior del carromato de provisiones junto a los civiles que lloraban desconsolados. Me puso al cuello el medallón de bronce con el león grabado de la familia.",
        "—Escúchame bien, Seier —me susurró al oído, plantando un beso en mi frente mientras sus dedos manchados de hollín acariciaban mi mejilla—. Tienes un alma más sabia de lo que aparentas. Lo he visto en tus ojos desde que eras un bebé. Vive. No permitas que el orgullo estúpido te mate. Hazte fuerte... más fuerte que cualquier demonio, y más implacable que cualquier noble traidor.",
        "Fue la última vez que los vi.",
        "Esa noche, Ronald y Elena Silva, junto a un puñado de cuarenta soldados leales, defendieron el cuello de botella del desfiladero contra más de ochocientos engendros demoníacos. Cuando la vanguardia del Rey finalmente llegó tres días después, encontraron los cuerpos de mis padres erguidos sobre una montaña de cadáveres putrefactos, cubiertos de heridas que habrían derribado a diez gigantes.",
        "Los civiles se salvaron. El bastión cayó, pero la invasión fue frenada.",
        "* * *",
        "El Rey del Reino de Oakhaven aplastó los remanentes demoníacos en las fronteras y, para calmar el clamor popular sin enemistarse abiertamente con los nobles que orquestaron la traición, emitió un edicto real con sello de cera dorada:",
        "La Casa Silva recibiría honores póstumos de héroes defensores. Y su único vástago, el huérfano Seier Silva, tendría garantizada una plaza irrevocable en la prestigiosa Gran Academia Militar y Mágica de la Capital cuando alcanzara los dieciséis años de edad, libre de costes y con todos los derechos de un noble de sangre.",
        "Un trozo de pergamino adornado a cambio de dos vidas.",
        "Acepté el pergamino sin derramar una sola lágrima.",
        "No me encerré en la amargura ni me dediqué a lamentar mi suerte en un rincón. Si algo me había enseñado mi vida anterior en la Tierra, era que el mundo no le debe nada a nadie. La debilidad engendra tragedia; el poder, en cambio, dicta las reglas.",
        "Durante los diez años siguientes, viví en los límites del antiguo feudo, bajo el cuidado de viejos campesinos y veteranos mutilados que veneraban la memoria de mi padre. Y mientras los niños nobles de la capital aprendían etiqueta, bailes de salón y conjuros vistosos para impresionar a damiselas, yo convertí mi cuerpo en una máquina de guerra.",
        "Recordaba los entrenamientos de mi vida pasada: las tardes agotadoras en el gimnasio levantando hierro pesado, el sonido seco de los guantes golpeando el saco y las enseñanzas del Muay Thai y el Jeet Kune Do.",
        "Golpes directos. Sin acrobacias inútiles ni movimientos de exhibición. Condicionamiento óseo: golpeé troncos de abedul con las tibias hasta que el periostio se calcificó y las piernas se volvieron tan duras como barras de hierro. Trabajé el agarre del cuello, los rodillazos demoledores en la corta distancia y los codos que cortaban carne y hueso con la precisión de un bisturí.",
        "Y lo más importante: aprendí a fusionar ese estilo callejero y pragmático con el maná de este mundo.",
        "No tenía talento para disparar bolas de fuego a cincuenta metros ni para invocar vientos huracanados. En su lugar, aprendí a comprimir la energía mágica hacia adentro: dentro de mis fibras musculares, mis tendones y mi densidad ósea. Cada golpe no era solo fuerza bruta; era una detonación de masa compacta.",
        "* * *",
        "El día de mi cumpleaños número dieciséis, me miré en el reflejo de un balde con agua.",
        "No era el típico protagonista de cuentos de hadas: medía un metro con setenta y cinco y pesaba noventa kilos de músculo macizo y denso. Mi cuello era ancho, mis hombros parecían dos bloques de granito y mis nudillos estaban cubiertos de callos oscuros. Mi cabello castaño estaba algo revuelto y mi expresión solía parecer tosca, casi indiferente, fruto de una década de aislamiento y esfuerzo.",
        "No era rápido como una liebre ni saltaba tres metros en el aire, pero sabía que si lograba cerrar la distancia y poner las manos encima de un enemigo, nada en este mundo se mantendría en pie.",
        "Me colgué al hombro una gastada mochila de cuero, me ajusté el abrigo sobre la camisa de lino y guardé el pergamino sellado del Rey junto a mi pecho.",
        "—Es hora —me dije a mí mismo con voz grave y áspera.",
        "El camino hacia la Capital y la Academia había comenzado. Y quienes creyeran que la Casa Silva estaba muerta y enterrada... pronto iban a descubrir cuán equivocados estaban."
    ]
}

CAPITULO_1 = {
    "header": "Capítulo 1",
    "title": "El Bruto y la Dama de Escarcha",
    "paragraphs": [
        "La Gran Capital de Solaria era un monumento al exceso y la arrogancia.",
        "Enormes murallas de mármol blanco se alzaban hacia el cielo, coronadas por agujas de cristal mágico que refractaban la luz solar en destellos dorados y violetas. Las calles principales estaban adoquinadas con precisión geométrica, transitadas por carruajes de maderas nobles tirados por corceles con arneses de plata.",
        "Hombres con jubones de seda perfumada y mujeres con corsés ceñidos y miriñaques paseaban bajo sombrillas, mirándome como si una bestia salvaje se hubiera colado en su jardín privado.",
        "Y tenían motivos.",
        "Con mi metro setenta y cinco de estatura y noventa kilos de puro músculo concentrado, mi paso era pesado y constante. Mi abrigo de cuero desgastado y mis botas con restos de barro de camino contrastaban salvajemente con la pulcritud de la aristocracia urbana. No me importaba en lo más mínimo. El hambre que sentía tras cuatro días de viaje a lomo de carreta no se saciaba con miradas de desprecio.",
        "Divisé un cartel de roble colgando en una esquina: *La Jarra Plateada*, una posada de aspecto acomodado cerca del distrito militar de la Academia. Empujé la puerta y entré.",
        "* * *",
        "El local estaba repleto de aspirantes y estudiantes de la Academia. Se respiraba una mezcla de perfumes caros, cerveza especiada y la tensión previa al examen de asignación del día siguiente.",
        "Me acerqué al mostrador de pino.",
        "—Tres raciones completas de carne asada, una hogaza entera de pan de centeno y una jarra grande de agua fresca —dije con mi voz profunda, apoyando unas cuantas monedas de plata sobre la madera.",
        "La mesera, una muchacha joven con orejas redondeadas de humana, parpadeó sorprendida por el tamaño del pedido, pero al ver el brillo de la plata asintió con rapidez y corrió a la cocina.",
        "Me senté en una mesa vacía del rincón, apoyé la espalda contra la pared y crucé los brazos sobre el pecho. La tela de mi camisa tiraba bajo la tensión de mis hombros. Mientras esperaba, me dediqué a observar.",
        "Había jóvenes nobles de todas las calañas: espadachines con roperas delgadas como agujas, magos portando báculos coronados con gemas elementales, y miembros de razas semi-humanas y élficas pertenecientes a embajadas aliadas. La mayoría se movía con poses teatrales, discutiendo sobre linajes, duelos de honor y círculos de magia de tercer orden.",
        "Florituras. Movimientos innecesarios. En el barro de la frontera, ese tipo de arrogancia te costaba la garganta antes del primer suspiro.",
        "Mi plato llegó humeante. Agarré el primer trozo de carne con los dedos y le di un mordisco profundo. Los jugos y la sal me devolvieron la energía al instante. Apenas iba por la mitad cuando el murmullo de la taberna se interrumpió por una discusión estridente cerca de la entrada.",
        "—¡Te he dicho que te apartes, campesina! ¿Acaso no ves con quién estás hablando? —rugió una voz chillona y altiva.",
        "Alcé la vista mientras masticaba con calma.",
        "Un trío de jóvenes nobles, enfundados en elegantes capas azul marino con el emblema de la espada cruzada de la Academia, acorralaban a la misma mesera que me había servido. El líder, un muchacho rubio de cabello engominado y una ropera con empuñadura dorada al cinto, la sostenía del brazo con fuerza innecesaria.",
        "—Has derramado vino sobre la capa de un Barón de la Casa Montfort —continuó el rubio, empujándola contra una mesa—. Te va a costar el sueldo de todo un mes... o podrías subir a mi habitación y compensarme de otra forma más placentera.",
        "Sus dos acompañantes rieron con suficiencia. Varios clientes miraron hacia otro lado, intimidados por el apellido noble.",
        "En medio del forcejeo, uno de los lacayos retrocedió sin mirar y chocó con violencia contra mi mesa, haciendo que mi jarra de agua y la mitad de mi hogaza de pan cayeran al suelo sucio.",
        "Dejé el trozo de carne en el plato.",
        "El lacayo me miró con fastidio, barriendo mi ropa con desdén.",
        "—¿Qué miras, mastodonte? Límpialo tú y quita esa cara de imbécil si no quieres que te hagamos azotar —escupió el noble menor.",
        "No respondí de inmediato. Me puse de pie despacio.",
        "Al erguirme por completo, mi sombra cubrió casi toda la mesa. Mis noventa kilos no eran grasa flácida; la silueta de mi espalda en V y el grosor de mis brazos se marcaron bajo el cuero. El chico tuvo que dar un paso involuntario hacia atrás, tragando saliva.",
        "—Estás pisando mi comida —dije en un tono plano, sin alzar la voz, pero con la resonancia de un trueno lejano.",
        "—¿Qué... qué has dicho, perro sarnoso? —el líder rubio soltó a la mesera y dio un paso al frente, tocando la empuñadura de su espada—. ¿Sabes con quién hablas? ¡Soy Cedric de Montfort! ¡Deberías estar de rodillas agradeciendo que respire tu mismo aire!",
        "Cedric desenvainó su espada con un silbido metálico. El filo destelló con una runa de maná de viento que silbaba en el aire. Con una pose de esgrima clásica, adelantó el pie y lanzó una estocada rápida y punzante directo hacia mi hombro, buscando ensartarme para enseñarme una lección.",
        "Fue un error garrafal.",
        "No retrocedí ni un milímetro. Apliqué la máxima del Jeet Kune Do: *interceptar el ataque en su origen*.",
        "Con un movimiento seco y milimétrico, desvié la trayectoria de su hoja usando el dorso de mi antebrazo izquierdo reforzado con una pulsación invisible de maná concentrado. El acero sonó como si hubiera golpeado una viga de acero.",
        "Antes de que Cedric pudiera comprender por qué su espada se había desviado, ya había cerrado la distancia entrando en el espacio ciego de su guardia.",
        "Agarré con mi mano derecha su nuca con un agarre de clinch de Muay Thai inflexible. Su cabeza quedó atrapada.",
        "—¿Qué demo...?",
        "Subí mi rodilla derecha como un ariete hacia su plexo solar.",
        "*¡CRACK!*",
        "El impacto fue sordo, pesado y brutal. El aire abandonó los pulmones de Cedric en una bocanada agónica. Sus ojos se abrieron como platos y la ropera cayó al suelo con un tintineo metálico.",
        "Sin soltarlo, giré las caderas y le propiné un codazo corto con el filo del codo izquierdo directamente en el lateral de la mandíbula.",
        "Cedric cayó como un saco de patatas inconsciente sobre las tablas de madera, con un hilo de baba y sangre escurriéndose por la comisura de los labios.",
        "Todo ocurrió en menos de dos segundos.",
        "Los otros dos nobles se quedaron paralizados, con el rostro blanco como el papel. Uno de ellos intentó conjurar una llama en la palma de su mano, pero di medio paso al frente, agarré su muñeca, la torcí con una palanca simple y lo estampé de espaldas contra la barra del bar, quejándose de dolor.",
        "El tercero simplemente alzó las manos y dio tres pasos atrás, temblando.",
        "—La carne y el pan costaban tres monedas de cobre —dije con voz áspera, sacando la bolsa de monedas de Cedric de su cinturón, tomando exactamente tres monedas y arrojándole el resto sobre el pecho inconsciente—. El desperdicio no se tolera.",
        "La taberna entera estaba en silencio absoluto.",
        "* * *",
        "Me disponía a darme la vuelta para buscar otra mesa cuando sentí una presencia distinta. Un cosquilleo gélido en la nuca, como si una aguja de hielo estuviera apuntando entre mis omóplatos.",
        "Volteé la mirada hacia la zona más apartada de la sala, junto al ventanal con cortinas de terciopelo.",
        "Allí, sentada sola frente a una copa de licor cristalino, una joven me observaba fijamente.",
        "Era una elfa de una belleza deslumbrante y aristocrática, aunque desprovista de cualquier calidez. Su cabello era de un color plata pulida, lacio y cayendo sobre sus hombros como hilos de luna; sus orejas largas y afiladas sobresalían con gracia entre los mechones. Vestía un uniforme oficial de la Academia con bordados plateados y ribetes de escarcha que denotaban el rango de las familias más influyentes del imperio.",
        "Bajo la chaqueta de oficial ceñida, la curva de su busto se marcaba con firmeza sobre un corsé de cuero blanco, y sus piernas largas y esbeltas estaban enfundadas en calzas oscuras que remataban en botas altas de jinete.",
        "Pero lo que más llamaba la atención eran sus ojos: dos zafiros gélidos, afilados como dagas, que me examinaban con una mezcla de curiosidad analítica y un evidente desdén.",
        "Nuestras miradas se cruzaron durante varios segundos.",
        "Cualquier plebeyo o noble menor habría bajado los ojos ante esa presencia intimidante y noble. Yo no lo hice. Sostuve su mirada con tranquilidad, apreciando sin disimulo la firmeza de su porte y la tensión que emanaba de su figura.",
        "La elfa arqueó una ceja plateada. Dejó su copa sobre la mesa con un chasquido seco y se puso de pie. Medía casi un metro con setenta, de complexión grácil pero con la fibra de una espadachina de élite.",
        "Caminó hacia mí con pasos ligeros y silenciosos, deteniéndose a apenas dos pasos de distancia. El aire a su alrededor se sentía varios grados más frío, impregnado de un aroma tenue a lirios de nieve y menta.",
        "—Una técnica repugnante y carente de elegancia —dijo con una voz melodiosa pero cortante como el cristal roto—. Sin postura de espada, sin círculos de canalización y apoyándote en la violencia bruta de una bestia salvaje. Típico de los brutos que llegan de las provincias olvidadas.",
        "La miré de arriba abajo sin inmutarme.",
        "—Funcionó —respondí secamente—. Tu amigo el barón está durmiendo la mona en el suelo. Y tú estás demasiado cerca de alguien a quien consideras una bestia.",
        "Un destello de irritación brilló en el fondo de sus ojos azul zafiro. Pocos se atrevían a hablarle con semejante desparpajo.",
        "—No es mi amigo. Los Montfort son perros ruidosos que no merecen ni la espada que portan —replicó ella, cruzando los brazos bajo su pecho, lo que hizo que su figura se acentuara de manera inevitable frente a mis ojos—. Mi nombre es Elysea de Vane. Y mañana, en el Coliseo de Asignación de la Academia, verás la diferencia entre un verdadero combatiente del imperio y un campesino musculoso que juega a ser guerrero.",
        "—Seier —dije simplemente, sin dar títulos ni reverencias—. Y si mañana te toca cruzar la espada conmigo, Elysea... más te vale no pestañear. Mis rodillas no piden permiso a los apellidos nobles.",
        "La elfa apretó los labios con una mezcla de furia contenida y un brillo involuntario de desafío que encendió sus pupilas. Se dio media vuelta con un movimiento fluido de su capa y salió de la taberna, dejando tras de sí una estela de aire helado y el eco de sus tacones en el pavimento.",
        "Miré hacia la puerta por donde había desaparecido.",
        "Una sonrisa tenue, casi imperceptible, asomó por primera vez en la comisura de mis labios.",
        "La Academia prometía ser bastante más interesante de lo que había imaginado."
    ]
}

def compilar_novela():
    titulo = "El camino del reencarnado"
    subtitulo = "Crónicas de Sangre, Acero y Deseo"
    autor = "Seier Silva"
    
    capitulos = [PROLOGO, CAPITULO_1]
    
    directorio_base = os.path.dirname(__file__)
    salida_docx = os.path.join(directorio_base, "compilado", "El_Camino_del_Reencarnado.docx")
    
    # Guardar copias en texto plano y en Word (.docx) individual para cada capítulo
    os.makedirs(os.path.join(directorio_base, "capitulos"), exist_ok=True)
    
    with open(os.path.join(directorio_base, "capitulos", "00_Prologo.txt"), "w", encoding="utf-8") as f:
        f.write(f"=== {PROLOGO['header'].upper()}: {PROLOGO['title'].upper()} ===\n\n")
        f.write("\n\n".join(PROLOGO['paragraphs']))
    create_single_chapter_document(PROLOGO, os.path.join(directorio_base, "capitulos", "00_Prologo.docx"))
        
    with open(os.path.join(directorio_base, "capitulos", "01_Capitulo_1.txt"), "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_1['header'].upper()}: {CAPITULO_1['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_1['paragraphs']))
    create_single_chapter_document(CAPITULO_1, os.path.join(directorio_base, "capitulos", "01_Capitulo_1.docx"))

    try:
        create_novel_document(titulo, subtitulo, autor, capitulos, salida_docx)
    except PermissionError:
        print(f"[AVISO] '{salida_docx}' está abierto en Microsoft Word. Ciérralo si deseas regenerarlo.")



if __name__ == "__main__":
    compilar_novela()
