[Read this in English](README.en.md)

# Esfer@ Acta Extractor

Una eina en Python per extreure i processar els registres de qualificacions de les actes de notes en PDF d’Esfer@. L’eina extreu les taules dels PDF, processa les notes dels estudiants i genera un fitxer Excel amb una vista estructurada de les qualificacions de RA (Resultats d’Aprenentatge) i de les qualificacions de MP (Mòdul Professional).

## Característiques

- Extreu taules dels informes de les actes de notes en PDF  
- Processa els noms i les notes dels estudiants  
- Identifica els MP amb entrades EM (amb hores d'estada a l'empresa)  
- Genera fitxers Excel amb:  
  - Notes dels estudiants per a cada RA  
  - Columnes espaiades correctament per a les qualificacions de MP  
  - Disposició especial de columnes per als MP amb entrades EM  
  - Protecció de cèl·lules: Totes les cèl·lules estan bloquejades excepte les cèl·lules de percentatges de RA a la darrera fila. Utilitzeu la contrasenya "edita'm" per desbloquejar.  

## Requisits

- Docker (recomanat)  
- Python 3.9+ (si s’executa localment)  
- Paquets de Python necessaris enumerats a `requirements.txt`  

## Estructura del projecte

```
.
├── input_pdf_files/          # Directori per als fitxers PDF d'entrada
│   └── .gitkeep              # Manté el directori a git però ignora els continguts
├── output_xlsx_files/        # Directori per als fitxers Excel de sortida
│   └── .gitkeep              # Manté el directori a git però ignora els continguts
├── src/                      # Paquet de codi font
│   ├── __init__.py           # Inicialització del paquet i exports
│   ├── pdf_processor.py      # Extracció i processament de PDFs
│   ├── data_processor.py     # Utilitats de processament de dades generals
│   ├── grade_processor.py    # Lògica específica de qualificacions
│   └── excel_processor.py    # Generació i formatació d'Excel
├── Dockerfile                # Configuració del contenidor Docker
├── README.md                 # Aquest fitxer (català)
├── README.en.md              # Versió anglès
├── requirements.txt          # Dependències de Python
├── LICENSE                   # Llicència GNU GPL-3.0
└── esfera-acta-extractor.py  # Script principal

```

### Descripció dels mòduls

- **pdf_processor.py**: Gestiona totes les operacions relacionades amb PDF  
  - Extracció de taules dels PDF  
  - Extracció del codi de grup de la primera pàgina  

- **data_processor.py**: Utilitats generals de processament de dades  
  - Normalització d'encapçalaments  
  - Filtrat i transformació de columnes  
  - Neteja i reformat de dades  

- **grade_processor.py**: Lògica específica de qualificacions  
  - Extracció de registres de RA  
  - Identificació de codi MP  
  - Detecció d'entrades EM  
  - Ordenació de registres  

- **excel_processor.py**: Gestió de fitxers Excel  
  - Generació de fitxers Excel  
  - Espaiat i organització de columnes  
  - (Futur) Format i fórmules  

## Instal·lació i ús

### Ús amb Docker (recomanat)

**Important**: Col·loca el teu fitxer PDF d'Esfer@ a la carpeta arrel del projecte abans d'executar qualsevol comanda de Docker.

1. Construeix la imatge de Docker:
```bash
docker build -t esfera-acta-extractor .
```

Tens dues opcions per executar el contenidor:

   a. Execució directa:
   ```bash
   docker run -v $(pwd):/app esfera-acta-extractor python esfera-acta-extractor.py
   ```

   b. Mode interactiu (recomanat per depuració o múltiples fitxers):
   ```bash
   docker run --rm -it \
     -v "$(pwd)":/data \
     -w /data \
     esfera-acta-extractor
   ```
   Una vegada dins el contenidor, pots executar:
   ```bash
   python esfera-acta-extractor.py
   ```
   Per sortir de la shell interactiva, simplement escriu:
   ```bash
   exit
   ```

**Nota**: El script:
- Buscarà els fitxers PDF en el directori `input_pdf_files`
- Processarà cada fitxer PDF individualment
- Generarà els fitxers Excel de sortida en el directori `output_xlsx_files`
- Anomenarà cada fitxer de sortida basant-se en el codi de grup trobat en el seu PDF corresponent
- Saltarà els fitxers que no són PDFs o no es poden processar
- Continuarà processant els fitxers restants fins i tot si algun falla

3. Neteja del contenidor:
   - Per al mode interactiu (opció --rm): el contenidor s'elimina automàticament en sortir
   - Per a processos en segon pla: atura el contenidor amb:
   ```bash
   docker stop $(docker ps -q --filter ancestor=esfera-acta-extractor)
   ```

**Nota**: L'eina:
- Buscarà el fitxer PDF al directori actual
- Generarà el fitxer Excel de sortida al mateix directori
- Anomenarà el fitxer de sortida segons el codi de grup trobat al PDF

#### Instal·lació local

1. Clona el repositori:
```bash
git clone https://github.com/Marcos-A/esfera-acta-extractor.git
cd esfera-acta-extractor
```

2. Instal·la les dependències:
```bash
pip install -r requirements.txt
```

3. Executa l'_script_:
```bash
python esfera-acta-extractor.py
```

## Entrada/Sortida

### Entrada
- Fitxer PDF d'Esfer@ amb taules de qualificacions
- El fitxer ha d'incloure el camp "Codi del grup" a la primera pàgina
- Formats de qualificació esperats: A#, PDT, EP, NA

### Sortida
- Fitxer Excel anomenat amb el codi de grup (p. ex., `CFPM_AG10101.xlsx`)
- Conté:
  - Noms dels estudiants
  - Notes de RA agrupades per MP
  - 3 columnes buides després dels MP amb entrades EM (etiquetades com a CENTRE, EMPRESA, MP)
  - 1 columna buida després dels MP normals (etiquetada com a MP)

## Desenvolupament

El projecte utilitza:
- `pandas` per a la manipulació de dades
- `pdfplumber` per al processament de PDFs
- `openpyxl` per a la generació de fitxers Excel
- `tabulate` per a la sortida de depuració/desenvolupament

## Contribuir

1. Feu un _fork_ del repositori
2. Creeu la teva branca de funció (`git checkout -b feature/AmazingFeature`)
3. Feu _commit_ dels vostres canvis (`git commit -m 'Add some AmazingFeature'`)
4. Pugeu la branca (`git push origin feature/AmazingFeature`)
5. Obriu una _Pull Request_

## Llicència

Aquest projecte està llicenciat sota la GNU General Public License v3.0 – consulteu el fitxer [LICENSE](LICENSE) per a més detalls.

Això significa que podeu:
- Utilitzar el programari per a qualsevol propòsit
- Modificar el programari per adaptar-lo a les teves necessitats
- Compartir el programari amb altra gent
- Compartir els canvis que en facis

Però heu de:
- Compartir el codi font quan comparteixis el programari
- Llicenciar qualsevol obra derivada sota GPL-3.0
- Indicar canvis significatius realitzats al programari
- Incloure la llicència original i els avisos de _copyright_

## Agraïments

- Creat per processar els informes de qualificacions educatives d'Esfer@
- Dissenyat per gestionar formats de PDF específics del sistema educatiu català
