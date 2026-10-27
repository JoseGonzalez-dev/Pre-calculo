# Escalamiento y Servidores de un Sistema en la Nube

Proyecto final de **Precálculo (0900-006)** · Escuela de Ingeniería en Sistemas · Universidad Mariano Gálvez de Guatemala.

Modelamos una plataforma en la nube con tres tipos de funciones para proyectar su crecimiento, encontrar su punto de saturación y optimizar su rentabilidad. Combina los proyectos 3 (rendimiento de un servidor) y 5 (crecimiento de usuarios) del Banco de Ideas.

## Modelos

| Componente | Tipo de función | Modelo | Variables |
|---|---|---|---|
| Crecimiento de usuarios | Exponencial | `U(t) = 150 · (1.25)^t` | `t`: meses |
| Latencia del servidor | Racional | `L(u) = 100 / (5000 − u)` | `u`: usuarios activos, `L` en ms. Asíntota vertical en `u = 5000` (colapso) |
| Utilidad financiera | Cuadrática | `P(x) = −0.5x² + 1000x − 5000` | `x`: precio de suscripción ($) |

## Resultados

| Pregunta | Desarrollo | Resultado |
|---|---|---|
| **1.** ¿En qué mes se superan 3,000 usuarios? | `150·(1.25)^t = 3000` → `t = ln(20) / ln(1.25)` | `t ≈ 13.43` → **mes 14** (`U(13) ≈ 2,728`, `U(14) ≈ 3,411`) |
| **2.** ¿Desde cuántos usuarios la latencia supera 2,000 ms? | `100/(5000 − u) > 2000` → `u > 5000 − 0.05` | **`u > 4,999.95` usuarios** |
| **3.** ¿Qué precio maximiza la utilidad? | Vértice: `h = −b/(2a) = 1000`, `k = P(1000)` | **Precio $1,000, utilidad máxima $495,000** |

**Para las conclusiones**
- Con este modelo la latencia solo se dispara pegada a la asíntota: el servidor responde bien hasta casi el final y luego colapsa de golpe, sin degradación gradual.
- `U(t)` alcanza los 5,000 usuarios en `t ≈ 15.71` meses (mes 16). En el mes 15 hay unos 4,263 usuarios, por lo que el escalamiento debe hacerse **antes del mes 15**.

## Requisitos

- Python 3.10 o superior
- `numpy`, `matplotlib` y `openpyxl` (ver `requirements.txt`)
- `tkinter`, que matplotlib usa para abrir la ventana interactiva. En Windows y macOS ya viene incluido con el instalador de Python.

## Instalación en Windows

1. **Instalar Python** desde [python.org/downloads](https://www.python.org/downloads/). En la primera pantalla del instalador, marcar **"Add python.exe to PATH"** antes de darle a *Install Now*. Para comprobarlo, abrir PowerShell y ejecutar:

   ```powershell
   python --version
   ```

2. **Descargar el proyecto.** Con Git:

   ```powershell
   git clone https://github.com/JoseGonzalez-dev/Pre-calculo.git
   cd Pre-calculo
   ```

   O desde GitHub con **Code → Download ZIP**, descomprimir y abrir PowerShell dentro de la carpeta (clic derecho en la carpeta → *Abrir en Terminal*).

3. **Crear y activar el entorno virtual:**

   ```powershell
   python -m venv venv
   venv\Scripts\Activate.ps1
   ```

   Si PowerShell muestra un error de *"la ejecución de scripts está deshabilitada"*, se resuelve de una de estas dos formas:

   ```powershell
   # Opción A: permitir scripts solo para tu usuario (se hace una sola vez)
   Set-ExecutionPolicy -Scope CurrentUser RemoteSigned

   # Opción B: usar el CMD de Windows en lugar de PowerShell
   venv\Scripts\activate.bat
   ```

   Si el entorno quedó activo, el prompt empieza con `(venv)`.

4. **Instalar las dependencias:**

   ```powershell
   pip install -r requirements.txt
   ```

5. **Ejecutar:**

   ```powershell
   python proyecto.py
   ```

Cada vez que abras una terminal nueva, hay que volver a activar el entorno con `venv\Scripts\Activate.ps1` antes de ejecutar.

## Instalación en Linux / macOS

```bash
git clone https://github.com/JoseGonzalez-dev/Pre-calculo.git
cd Pre-calculo

python3 -m venv venv
source venv/bin/activate

pip install -r requirements.txt
python3 proyecto.py
```

En Linux, `tkinter` suele venir como paquete aparte:

```bash
# Fedora
sudo dnf install python3-tkinter

# Debian / Ubuntu
sudo apt install python3-tk
```

## Uso

**Consola + gráficas interactivas**

```bash
python proyecto.py
```

Imprime el desarrollo algebraico paso a paso de las tres preguntas y abre una ventana con las tres gráficas y cinco sliders:

| Slider | Qué cambia |
|---|---|
| Crecimiento mensual | La tasa `r` de la función exponencial |
| Usuarios iniciales | `U0` de la función exponencial |
| Capacidad (asíntota) | La asíntota vertical y la curva de latencia |
| Mes `t` | Un punto en la gráfica 1 y, con esos usuarios, su latencia en la gráfica 2 |
| Precio `x` | Un punto sobre la parábola de la gráfica 3 y su utilidad |

El botón de disquete de la barra de la ventana guarda las gráficas como imagen para el documento.

**Tablas en Excel**

```bash
python proyecto.py --excel
```

Genera `tablas_proyecto.xlsx` con las hojas `Parámetros`, `Resultados`, `Usuarios`, `Latencia` y `Utilidad`. Todo usa fórmulas ligadas a la hoja `Parámetros`: si cambias un valor en las celdas azules, las tablas y los resultados se recalculan. La hoja `Latencia` usa puntos fijos de `u`, así que si cambias la capacidad hay que ajustar esos puntos a mano.

## Estructura del repositorio

```
.
├── proyecto.py            # cálculos, gráficas interactivas y exportación a Excel
├── tablas_proyecto.xlsx   # tablas de valores ya generadas (se regenera con --excel)
├── requirements.txt       # dependencias
├── README.md
├── LICENSE
└── .gitignore
```

## Notas

- La gráfica de latencia usa escala logarítmica en el eje y. Sin ella, la curva se ve pegada al piso y luego sube como una pared, y el punto de 2,000 ms no se aprecia.
- La cuadrática usa los coeficientes del enunciado, que no se modifican con los sliders.
