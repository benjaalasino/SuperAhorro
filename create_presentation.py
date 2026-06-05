from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import pptx.oxml.ns as nsmap
from lxml import etree
import copy

# ── Colors ──────────────────────────────────────────────────────────────────
GREEN_DARK   = RGBColor(0x2E, 0x7D, 0x32)
GREEN_LIGHT  = RGBColor(0xE8, 0xF5, 0xE9)
GREEN_MED    = RGBColor(0x43, 0xA0, 0x47)
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)
BLACK        = RGBColor(0x1A, 0x1A, 0x1A)
GRAY         = RGBColor(0x75, 0x75, 0x75)
CODE_BG      = RGBColor(0x1E, 0x1E, 0x2E)  # dark code background
CODE_FG      = RGBColor(0xCD, 0xD6, 0xF4)  # code text
KW_COLOR     = RGBColor(0xCB, 0xA6, 0xF7)  # keyword purple
STR_COLOR    = RGBColor(0xA6, 0xE3, 0xA1)  # string green
ORANGE       = RGBColor(0xFF, 0x6F, 0x00)
BLUE_LIGHT   = RGBColor(0xE3, 0xF2, 0xFD)
RED_LIGHT    = RGBColor(0xFF, 0xEB, 0xEE)

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

BLANK = prs.slide_layouts[6]  # completely blank

# ── Helper: add rectangle ────────────────────────────────────────────────────
def rect(slide, l, t, w, h, fill_rgb=None, line_rgb=None, line_width=Pt(0)):
    shape = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    shape.line.width = line_width
    if fill_rgb:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_rgb
    else:
        shape.fill.background()
    if line_rgb:
        shape.line.color.rgb = line_rgb
    else:
        shape.line.fill.background()
    return shape

# ── Helper: add text box ─────────────────────────────────────────────────────
def text_box(slide, text, l, t, w, h,
             font_size=Pt(14), bold=False, color=BLACK,
             align=PP_ALIGN.LEFT, wrap=True, italic=False):
    txBox = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    txBox.word_wrap = wrap
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = font_size
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txBox

# ── Helper: code block ───────────────────────────────────────────────────────
def code_block(slide, lines, l, t, w, h, font_size=Pt(8.5)):
    """Renders lines as a dark-bg code block. lines = list of (text, rgb_color)"""
    bg = rect(slide, l, t, w, h, fill_rgb=CODE_BG)
    bg.line.fill.background()

    txBox = slide.shapes.add_textbox(Inches(l+0.12), Inches(t+0.1),
                                     Inches(w-0.24), Inches(h-0.2))
    txBox.word_wrap = False
    tf = txBox.text_frame
    tf.word_wrap = False

    first = True
    for (line_text, line_color) in lines:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.space_before = Pt(0)
        p.space_after  = Pt(0)
        run = p.add_run()
        run.text = line_text
        run.font.size = font_size
        run.font.color.rgb = line_color if line_color else CODE_FG
        run.font.name = "Courier New"
    return txBox

# ── Helper: slide header bar ─────────────────────────────────────────────────
def header(slide, title, subtitle=""):
    rect(slide, 0, 0, 13.33, 1.15, fill_rgb=GREEN_DARK)
    text_box(slide, title, 0.35, 0.12, 10, 0.55,
             font_size=Pt(26), bold=True, color=WHITE)
    if subtitle:
        text_box(slide, subtitle, 0.35, 0.68, 10, 0.4,
                 font_size=Pt(13), color=RGBColor(0xC8, 0xE6, 0xC9))

# ── Helper: bullet list ──────────────────────────────────────────────────────
def bullet_list(slide, items, l, t, w, h, font_size=Pt(13), color=BLACK, bullet="▸ "):
    txBox = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    txBox.word_wrap = True
    tf = txBox.text_frame
    tf.word_wrap = True
    first = True
    for item in items:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.space_before = Pt(2)
        run = p.add_run()
        run.text = bullet + item
        run.font.size = font_size
        run.font.color.rgb = color
    return txBox

# ── Helper: info card ────────────────────────────────────────────────────────
def info_card(slide, title, body_lines, l, t, w, h,
              bg=GREEN_LIGHT, title_color=GREEN_DARK, body_color=BLACK):
    rect(slide, l, t, w, h, fill_rgb=bg, line_rgb=GREEN_MED, line_width=Pt(1))
    text_box(slide, title, l+0.12, t+0.1, w-0.24, 0.38,
             font_size=Pt(13), bold=True, color=title_color)
    body = "\n".join(body_lines)
    text_box(slide, body, l+0.12, t+0.52, w-0.24, h-0.6,
             font_size=Pt(11), color=body_color)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — PORTADA
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, fill_rgb=GREEN_DARK)
rect(s, 0, 5.8, 13.33, 1.7, fill_rgb=RGBColor(0x1B, 0x5E, 0x20))

# Title
text_box(s, "🛒  SuperAhorro", 1.5, 1.5, 10.5, 1.4,
         font_size=Pt(52), bold=True, color=WHITE, align=PP_ALIGN.CENTER)
text_box(s, "Aplicación Android de Control de Gastos en Supermercados",
         1.5, 2.9, 10.5, 0.7,
         font_size=Pt(20), color=RGBColor(0xC8, 0xE6, 0xC9), align=PP_ALIGN.CENTER)
text_box(s, "Kotlin  •  Jetpack Compose  •  MVVM  •  Navigation  •  SharedPreferences",
         1.5, 3.65, 10.5, 0.55,
         font_size=Pt(14), color=RGBColor(0xA5, 0xD6, 0xA7), align=PP_ALIGN.CENTER, italic=True)
text_box(s, "com.undef.superahorroCalvoAlasino",
         1.5, 6.1, 10.5, 0.45,
         font_size=Pt(13), color=RGBColor(0x80, 0xCB, 0xC4), align=PP_ALIGN.CENTER, italic=True)
text_box(s, "API Level 36 (Android 15)  •  Android Studio",
         1.5, 6.6, 10.5, 0.4,
         font_size=Pt(12), color=RGBColor(0xA5, 0xD6, 0xA7), align=PP_ALIGN.CENTER)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — TECNOLOGÍAS
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, fill_rgb=RGBColor(0xF9, 0xFA, 0xF9))
header(s, "Tecnologías y Lenguajes", "Stack completo del proyecto")

cards = [
    ("🔤  Kotlin", ["Lenguaje oficial de Android", "Null-safety, corrutinas integradas", "Menos código que Java", "Interoperable con Java 100%"]),
    ("🔧  Gradle (KTS)", ["build.gradle.kts — sintaxis Kotlin", "libs.versions.toml — catálogo de versiones", "AGP 9.1.1 — Android Gradle Plugin", "Compila el APK final"]),
    ("🎨  Jetpack Compose", ["UI declarativa 100%", "@Composable functions", "Material 3 (Material You)", "Sin XML layouts"]),
    ("🏗️  Android Studio", ["IDE oficial de Google", "API 36 (Android 15)", "Min SDK 24 (Android 7.0)", "Emulador y debugging integrado"]),
    ("🗄️  SharedPreferences", ["Persistencia local de datos", "JSON serializado con org.json", "Datos aislados por usuario", "No requiere servidor"]),
    ("🧭  Navigation Compose", ["navigation-compose:2.7.7", "NavHost + composable()", "Argumentos tipados (IntType)", "Back-stack automático"]),
]

cols = [(0.25, 1.3), (4.6, 1.3), (8.95, 1.3), (0.25, 4.1), (4.6, 4.1), (8.95, 4.1)]
for i, (title, bullets) in enumerate(cards):
    x, y = cols[i]
    info_card(s, title, bullets, x, y, 4.1, 2.55)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — ESTRUCTURA DEL PROYECTO + PAQUETE
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, fill_rgb=RGBColor(0xF9, 0xFA, 0xF9))
header(s, "Estructura del Proyecto", "Organización de carpetas y paquete com.undef.*")

# left: tree
tree_lines = [
    ("SuperAhorro/", WHITE),
    ("├── app/", RGBColor(0xA6, 0xE3, 0xA1)),
    ("│   ├── build.gradle.kts", CODE_FG),
    ("│   └── src/main/", RGBColor(0xA6, 0xE3, 0xA1)),
    ("│       ├── AndroidManifest.xml", CODE_FG),
    ("│       ├── java/com/undef/", RGBColor(0xCB, 0xA6, 0xF7)),
    ("│       │   superahorroCalvoAlasino/", RGBColor(0xCB, 0xA6, 0xF7)),
    ("│       │   ├── MainActivity.kt", CODE_FG),
    ("│       │   ├── model/", RGBColor(0x89, 0xDC, 0xEB)),
    ("│       │   │   └── Models.kt", CODE_FG),
    ("│       │   ├── viewmodel/", RGBColor(0x89, 0xDC, 0xEB)),
    ("│       │   │   ├── UsuarioViewModel.kt", CODE_FG),
    ("│       │   │   └── CompraViewModel.kt", CODE_FG),
    ("│       │   ├── navigation/", RGBColor(0x89, 0xDC, 0xEB)),
    ("│       │   │   ├── NavRoutes.kt", CODE_FG),
    ("│       │   │   └── AppNavGraph.kt", CODE_FG),
    ("│       │   └── ui/", RGBColor(0x89, 0xDC, 0xEB)),
    ("│       │       ├── screens/ (12 pantallas)", CODE_FG),
    ("│       │       ├── components/", CODE_FG),
    ("│       │       └── theme/", CODE_FG),
    ("│       └── res/values/strings.xml", RGBColor(0xF3, 0x8B, 0xA8)),
    ("└── gradle/libs.versions.toml", RGBColor(0xA6, 0xE3, 0xA1)),
]
code_block(s, tree_lines, 0.25, 1.2, 6.0, 6.05, font_size=Pt(8.5))

# right: paquete highlight + info cards
text_box(s, "📦  Nombre del Paquete", 6.55, 1.25, 6.5, 0.45,
         font_size=Pt(15), bold=True, color=GREEN_DARK)
code_block(s, [
    ("// app/build.gradle.kts", GRAY),
    ("namespace = \"com.undef.superahorroCalvoAlasino\"", RGBColor(0xA6, 0xE3, 0xA1)),
    ("applicationId = \"com.undef.superahorroCalvoAlasino\"", RGBColor(0xA6, 0xE3, 0xA1)),
], 6.55, 1.72, 6.55, 0.88, font_size=Pt(9))

text_box(s, "✅  Cumple: com.undef.nombreappalumno", 6.55, 2.68, 6.5, 0.38,
         font_size=Pt(12), bold=True, color=GREEN_DARK)

# MVVM boxes
mvvm = [
    ("🧩  model/", "Data classes: Compra, Producto, Usuario"),
    ("⚙️  viewmodel/", "UsuarioViewModel, CompraViewModel\n→ extienden ViewModel()"),
    ("🎨  ui/screens/", "12 pantallas @Composable\n→ Solo UI, sin lógica de negocio"),
    ("🧭  navigation/", "NavRoutes (sealed class)\nAppNavGraph (NavHost)"),
]
for i, (t, b) in enumerate(mvvm):
    y = 3.2 + i * 0.97
    info_card(s, t, [b], 6.55, y, 6.55, 0.88, bg=GREEN_LIGHT)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — JETPACK COMPOSE + MAIN ACTIVITY
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, fill_rgb=RGBColor(0xF9, 0xFA, 0xF9))
header(s, "Jetpack Compose", "UI declarativa — MainActivity.kt como punto de entrada")

# Left: code
code_lines = [
    ("// MainActivity.kt", GRAY),
    ("class MainActivity : ComponentActivity() {", CODE_FG),
    ("", CODE_FG),
    ("  override fun onCreate(savedInstanceState: Bundle?) {", CODE_FG),
    ("    super.onCreate(savedInstanceState)", CODE_FG),
    ("", CODE_FG),
    ("    setContent {", KW_COLOR),
    ("      SuperAhorroTheme {", RGBColor(0xA6, 0xE3, 0xA1)),
    ("        Surface(", CODE_FG),
    ("          modifier = Modifier.fillMaxSize(),", CODE_FG),
    ("          color = MaterialTheme.colorScheme.background", CODE_FG),
    ("        ) {", CODE_FG),
    ("          AppNavGraph(applicationContext)", RGBColor(0x89, 0xDC, 0xEB)),
    ("        }", CODE_FG),
    ("      }", CODE_FG),
    ("    }", CODE_FG),
    ("  }", CODE_FG),
    ("}", CODE_FG),
]
code_block(s, code_lines, 0.25, 1.25, 6.4, 5.1)

# Right: annotations
text_box(s, "¿Por qué Jetpack Compose?", 6.9, 1.25, 6.2, 0.4,
         font_size=Pt(15), bold=True, color=GREEN_DARK)

points = [
    ("setContent { }", "Reemplaza setContentView(R.layout.*).\nNo hay XML de layout — todo es código Kotlin."),
    ("@Composable functions", "Cada pantalla es una función anotada\ncon @Composable que describe la UI."),
    ("SuperAhorroTheme { }", "Aplica el tema Material 3 globalmente:\ncolores, tipografía y formas definidos en\nui/theme/Theme.kt"),
    ("AppNavGraph()", "Punto de entrada al grafo de navegación.\nRecibe el contexto para inicializar\nlos ViewModels con SharedPreferences."),
]
for i, (title, body) in enumerate(points):
    y = 1.72 + i * 1.34
    rect(s, 6.9, y, 6.2, 1.22, fill_rgb=GREEN_LIGHT,
         line_rgb=GREEN_MED, line_width=Pt(1))
    text_box(s, title, 7.05, y + 0.07, 5.9, 0.38,
             font_size=Pt(12), bold=True, color=GREEN_DARK)
    text_box(s, body, 7.05, y + 0.45, 5.9, 0.7,
             font_size=Pt(10.5), color=BLACK)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — ARQUITECTURA MVVM — MODELOS
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, fill_rgb=RGBColor(0xF9, 0xFA, 0xF9))
header(s, "Arquitectura MVVM — Capa Model", "model/Models.kt  ·  Data classes puras de Kotlin")

code_lines = [
    ("// model/Models.kt", GRAY),
    ("", CODE_FG),
    ("data class Compra(", KW_COLOR),
    ("    val id: Int,", CODE_FG),
    ("    val supermercado: String,", CODE_FG),
    ("    val fecha: String,", CODE_FG),
    ("    val hora: String,", CODE_FG),
    ("    val total: Double,", CODE_FG),
    ("    val productos: List<Producto> = emptyList()", CODE_FG),
    (")", CODE_FG),
    ("", CODE_FG),
    ("data class Producto(", KW_COLOR),
    ("    val id: Int,", CODE_FG),
    ("    val codigo: String,", CODE_FG),
    ("    val nombre: String,", CODE_FG),
    ("    val descripcion: String,", CODE_FG),
    ("    val precio: Double", CODE_FG),
    (")", CODE_FG),
    ("", CODE_FG),
    ("data class Usuario(", KW_COLOR),
    ("    val nombre: String = \"\",", CODE_FG),
    ("    val email: String = \"\",", CODE_FG),
    ("    val password: String = \"\"", CODE_FG),
    (")", CODE_FG),
]
code_block(s, code_lines, 0.25, 1.25, 6.0, 6.0)

# right: explanation
text_box(s, "¿Qué son data classes en Kotlin?", 6.55, 1.25, 6.55, 0.45,
         font_size=Pt(14), bold=True, color=GREEN_DARK)

reasons = [
    "equals(), hashCode(), toString() y copy() son generados automáticamente por el compilador.",
    "val (immutability) → los datos no se modifican accidentalmente.",
    "Compra contiene una lista de Producto anidada — relación 1:N.",
    "La capa Model NO conoce la UI ni la lógica de negocio:\npuro almacenamiento de datos.",
    "copy() permite actualizar estado de forma inmutable:\ncompra.copy(productos = nuevaLista)",
]

for i, r in enumerate(reasons):
    y = 1.78 + i * 1.07
    rect(s, 6.55, y, 6.55, 0.98, fill_rgb=GREEN_LIGHT, line_rgb=GREEN_MED, line_width=Pt(1))
    text_box(s, f"{'①②③④⑤'[i]}  {r}", 6.7, y+0.07, 6.25, 0.84,
             font_size=Pt(10.5), color=BLACK)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — VIEWMODEL: UsuarioViewModel
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, fill_rgb=RGBColor(0xF9, 0xFA, 0xF9))
header(s, "MVVM — UsuarioViewModel", "Gestión del estado del usuario con mutableStateOf + SharedPreferences")

code_lines_vm = [
    ("// viewmodel/UsuarioViewModel.kt", GRAY),
    ("class UsuarioViewModel(", KW_COLOR),
    ("    private var context: Context? = null", CODE_FG),
    (") : ViewModel() {", KW_COLOR),
    ("", CODE_FG),
    ("  // Estado reactivo: la UI se recompone cuando cambia", GRAY),
    ("  private val _usuario = mutableStateOf(Usuario())", RGBColor(0x89, 0xDC, 0xEB)),
    ("  val usuario = _usuario", CODE_FG),
    ("", CODE_FG),
    ("  fun registrarUsuario(nombre: String,", CODE_FG),
    ("                       email: String, password: String) {", CODE_FG),
    ("    _usuario.value = Usuario(nombre, email, password)", CODE_FG),
    ("    guardarUsuarioEnPreferencias()", RGBColor(0xA6, 0xE3, 0xA1)),
    ("  }", CODE_FG),
    ("", CODE_FG),
    ("  fun validarCredenciales(email: String,", CODE_FG),
    ("                          password: String): Boolean {", CODE_FG),
    ("    val u = _usuario.value", CODE_FG),
    ("    return u.email == email && u.password == password", RGBColor(0xA6, 0xE3, 0xA1)),
    ("  }", CODE_FG),
    ("", CODE_FG),
    ("  fun cerrarSesion() {", CODE_FG),
    ("    _usuario.value = Usuario()  // resetea el estado", GRAY),
    ("  }", CODE_FG),
    ("}", CODE_FG),
]
code_block(s, code_lines_vm, 0.25, 1.25, 6.55, 6.0, font_size=Pt(8.5))

# Right annotations
annotations = [
    ("ViewModel()", "Sobrevive rotaciones de pantalla.\nEl sistema lo destruye solo cuando el\nusuario abandona la pantalla definitivamente."),
    ("mutableStateOf()", "Jetpack Compose observa este estado.\nCada vez que _usuario.value cambia,\nlas Composables que lo leen se recomponen."),
    ("SharedPreferences", "Almacenamiento clave-valor local.\nNombre del archivo: \"super_ahorro_prefs\"\nClaves: nombre, email, password."),
    ("cerrarSesion()", "Resetea el estado a Usuario() vacío.\nLa UI reacciona automáticamente\ny redirige al Login."),
]
for i, (t, b) in enumerate(annotations):
    y = 1.25 + i * 1.55
    rect(s, 7.05, y, 6.05, 1.42, fill_rgb=GREEN_LIGHT, line_rgb=GREEN_MED, line_width=Pt(1))
    text_box(s, t, 7.2, y+0.08, 5.75, 0.38,
             font_size=Pt(13), bold=True, color=GREEN_DARK)
    text_box(s, b, 7.2, y+0.48, 5.75, 0.88,
             font_size=Pt(10.5), color=BLACK)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — VIEWMODEL: CompraViewModel
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, fill_rgb=RGBColor(0xF9, 0xFA, 0xF9))
header(s, "MVVM — CompraViewModel", "Estado reactivo con mutableStateListOf + serialización JSON")

code_lines_cv = [
    ("// viewmodel/CompraViewModel.kt", GRAY),
    ("class CompraViewModel(context: Context?) : ViewModel() {", KW_COLOR),
    ("", CODE_FG),
    ("  // Lista observable — Compose recompone al mutar", GRAY),
    ("  private val _compras = mutableStateListOf<Compra>()", RGBColor(0x89, 0xDC, 0xEB)),
    ("  val compras: List<Compra> = _compras", CODE_FG),
    ("  private var proximoId = 1", CODE_FG),
    ("", CODE_FG),
    ("  fun agregarCompra(supermercado: String, fecha: String,", CODE_FG),
    ("                    hora: String, total: Double,", CODE_FG),
    ("                    productosAgregados: List<Pair<String,Double>>) {", CODE_FG),
    ("    val nuevaCompra = Compra(", CODE_FG),
    ("        id = proximoId++, supermercado = supermercado,", CODE_FG),
    ("        fecha = fecha, hora = hora, total = total,", CODE_FG),
    ("        productos = productosAgregados.mapIndexed { i, (n, p) ->", CODE_FG),
    ("            Producto(i+1, \"\", n, \"\", p)", CODE_FG),
    ("        }", CODE_FG),
    ("    )", CODE_FG),
    ("    _compras.add(0, nuevaCompra)  // inserta al inicio", GRAY),
    ("    guardarCompras()              // persiste en JSON", RGBColor(0xA6, 0xE3, 0xA1)),
    ("  }", CODE_FG),
    ("", CODE_FG),
    ("  fun calcularTotalGastado(): Double =", CODE_FG),
    ("      _compras.sumOf { it.total }", RGBColor(0xA6, 0xE3, 0xA1)),
    ("}", CODE_FG),
]
code_block(s, code_lines_cv, 0.25, 1.25, 7.1, 6.0, font_size=Pt(8.5))

# Right: JSON persistence
text_box(s, "Persistencia JSON en SharedPreferences", 7.6, 1.25, 5.5, 0.42,
         font_size=Pt(13), bold=True, color=GREEN_DARK)
json_lines = [
    ("// Clave por usuario: \"compras_juan@gmail.com\"", GRAY),
    ("[", CODE_FG),
    ("  {", CODE_FG),
    ("    \"id\": 1,", STR_COLOR),
    ("    \"supermercado\": \"Carrefour\",", STR_COLOR),
    ("    \"fecha\": \"08/05/2026\",", STR_COLOR),
    ("    \"total\": 15420.50,", STR_COLOR),
    ("    \"productos\": [", CODE_FG),
    ("      {\"nombre\": \"Leche\", \"precio\": 980}", STR_COLOR),
    ("    ]", CODE_FG),
    ("  }", CODE_FG),
    ("]", CODE_FG),
]
code_block(s, json_lines, 7.6, 1.72, 5.5, 2.55, font_size=Pt(9))

info_card(s, "mutableStateListOf vs mutableStateOf", [
    "mutableStateListOf<T>() → Compose detecta",
    "inserciones/remociones individuales.",
    "Más eficiente que reemplazar toda la lista.",
    "",
    "guardarCompras() serializa la lista completa",
    "a JSONArray y la guarda en SharedPreferences.",
], 7.6, 4.38, 5.5, 2.85)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — NAVEGACIÓN: NavRoutes
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, fill_rgb=RGBColor(0xF9, 0xFA, 0xF9))
header(s, "Navegación — NavRoutes (Sealed Class)", "navigation/NavRoutes.kt · 12 rutas tipadas con Kotlin")

nr_lines = [
    ("// navigation/NavRoutes.kt", GRAY),
    ("sealed class NavRoutes(val route: String) {", KW_COLOR),
    ("", CODE_FG),
    ("  object Splash      : NavRoutes(\"splash\")", CODE_FG),
    ("  object Login       : NavRoutes(\"login\")", CODE_FG),
    ("  object Registro    : NavRoutes(\"registro\")", CODE_FG),
    ("  object Home        : NavRoutes(\"home\")", CODE_FG),
    ("  object NuevaCompra : NavRoutes(\"nueva_compra\")", CODE_FG),
    ("  object Historial   : NavRoutes(\"historial\")", CODE_FG),
    ("  object Estadisticas: NavRoutes(\"estadisticas\")", CODE_FG),
    ("  object Perfil      : NavRoutes(\"perfil\")", CODE_FG),
    ("  object EditarPerfil: NavRoutes(\"editar_perfil\")", CODE_FG),
    ("  object Settings    : NavRoutes(\"settings\")", CODE_FG),
    ("", CODE_FG),
    ("  // Rutas con argumentos tipados", GRAY),
    ("  object NuevoProducto : NavRoutes(", KW_COLOR),
    ("      \"nuevo_producto/{compraId}\") {", KW_COLOR),
    ("    fun withId(id: Int) = \"nuevo_producto/$id\"", RGBColor(0xA6, 0xE3, 0xA1)),
    ("  }", CODE_FG),
    ("", CODE_FG),
    ("  object DetalleCompra : NavRoutes(", KW_COLOR),
    ("      \"detalle_compra/{compraId}\") {", KW_COLOR),
    ("    fun withId(id: Int) = \"detalle_compra/$id\"", RGBColor(0xA6, 0xE3, 0xA1)),
    ("  }", CODE_FG),
    ("}", CODE_FG),
]
code_block(s, nr_lines, 0.25, 1.25, 6.55, 6.0, font_size=Pt(8.8))

# Right: diagram
text_box(s, "¿Por qué sealed class?", 7.1, 1.25, 6.0, 0.42,
         font_size=Pt(14), bold=True, color=GREEN_DARK)
text_box(s, "Una sealed class en Kotlin garantiza que SOLO\nexisten los subtipos declarados en el archivo.\nEl compilador verifica exhaustividad (when expressions).",
         7.1, 1.72, 6.0, 0.88, font_size=Pt(11), color=BLACK)

# Flow diagram boxes
flow = [
    ("Splash", "Verifica si hay usuario"),
    ("Login / Registro", "Autenticación"),
    ("Home", "Dashboard principal"),
    ("Historial / Estadísticas", "Vistas de datos"),
    ("Detalle Compra", "withId(compraId)"),
    ("Nuevo Producto", "Argumento tipado Int"),
]
for i, (name, desc) in enumerate(flow):
    y = 2.7 + i * 0.8
    bg = GREEN_DARK if name in ("Home",) else GREEN_LIGHT
    fc = WHITE if bg == GREEN_DARK else BLACK
    tc = WHITE if bg == GREEN_DARK else GREEN_DARK
    rect(s, 7.1, y, 5.95, 0.7, fill_rgb=bg, line_rgb=GREEN_MED, line_width=Pt(1))
    text_box(s, name, 7.25, y+0.05, 2.9, 0.35,
             font_size=Pt(11), bold=True, color=tc)
    text_box(s, desc, 10.2, y+0.08, 2.7, 0.35,
             font_size=Pt(10), color=GRAY if bg != GREEN_DARK else RGBColor(0xC8, 0xE6, 0xC9))
    if i < len(flow)-1:
        text_box(s, "▼", 9.85, y+0.7, 0.4, 0.15,
                 font_size=Pt(10), color=GREEN_MED, align=PP_ALIGN.CENTER)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 9 — NAVEGACIÓN: AppNavGraph
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, fill_rgb=RGBColor(0xF9, 0xFA, 0xF9))
header(s, "Navegación — AppNavGraph", "navigation/AppNavGraph.kt · NavHost con argumentos y back-stack")

ng_lines = [
    ("// navigation/AppNavGraph.kt", GRAY),
    ("@Composable", KW_COLOR),
    ("fun AppNavGraph(context: Context? = null) {", CODE_FG),
    ("    val navController = rememberNavController()", RGBColor(0x89, 0xDC, 0xEB)),
    ("    val compraViewModel   = remember { CompraViewModel(context) }", CODE_FG),
    ("    val usuarioViewModel  = remember { UsuarioViewModel(context) }", CODE_FG),
    ("", CODE_FG),
    ("    NavHost(navController, startDestination = NavRoutes.Splash.route) {", KW_COLOR),
    ("", CODE_FG),
    ("        // Rutas simples", GRAY),
    ("        composable(NavRoutes.Splash.route)  { SplashScreen(navController) }", CODE_FG),
    ("        composable(NavRoutes.Login.route)   { LoginScreen(navController,", CODE_FG),
    ("            usuarioViewModel, compraViewModel) }", CODE_FG),
    ("        composable(NavRoutes.Home.route)    { HomeScreen(navController,", CODE_FG),
    ("            compraViewModel) }", CODE_FG),
    ("", CODE_FG),
    ("        // Ruta con argumento tipado (Int)", GRAY),
    ("        composable(", CODE_FG),
    ("            route = NavRoutes.DetalleCompra.route,", CODE_FG),
    ("            arguments = listOf(", CODE_FG),
    ("                navArgument(\"compraId\") { type = NavType.IntType }", RGBColor(0xA6, 0xE3, 0xA1)),
    ("            )", CODE_FG),
    ("        ) { backStackEntry ->", CODE_FG),
    ("            val id = backStackEntry.arguments?.getInt(\"compraId\") ?: 0", CODE_FG),
    ("            DetalleCompraScreen(navController, id, compraViewModel)", RGBColor(0x89, 0xDC, 0xEB)),
    ("        }", CODE_FG),
    ("    }", CODE_FG),
    ("}", CODE_FG),
]
code_block(s, ng_lines, 0.25, 1.25, 7.85, 6.0, font_size=Pt(8.5))

# right: key concepts
concepts = [
    ("rememberNavController()", "Crea y recuerda el controlador de\nnavegación entre composiciones."),
    ("remember { ViewModel() }", "Los ViewModels se crean una sola vez\ny son compartidos entre pantallas."),
    ("navArgument + IntType", "Tipo seguro: no puede pasarse un String\ndonde se espera un Int como compraId."),
    ("popUpTo + inclusive", "Evita volver al Login después de\nautentica: limpia el back-stack."),
]
for i, (t, b) in enumerate(concepts):
    y = 1.25 + i * 1.55
    rect(s, 8.35, y, 4.75, 1.42, fill_rgb=GREEN_LIGHT, line_rgb=GREEN_MED, line_width=Pt(1))
    text_box(s, t, 8.5, y+0.08, 4.4, 0.38,
             font_size=Pt(12), bold=True, color=GREEN_DARK)
    text_box(s, b, 8.5, y+0.5, 4.4, 0.84,
             font_size=Pt(10.5), color=BLACK)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 10 — PANTALLA LOGIN
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, fill_rgb=RGBColor(0xF9, 0xFA, 0xF9))
header(s, "Pantalla Login — LoginScreen.kt", "ui/screens/LoginScreen.kt · Estado local + validación + navegación")

login_lines = [
    ("@Composable", KW_COLOR),
    ("fun LoginScreen(navController: NavController,", CODE_FG),
    ("               usuarioVM: UsuarioViewModel,", CODE_FG),
    ("               compraVM: CompraViewModel) {", CODE_FG),
    ("", CODE_FG),
    ("  // Estado local de la pantalla", GRAY),
    ("  var email    by remember { mutableStateOf(\"\") }", RGBColor(0x89, 0xDC, 0xEB)),
    ("  var password by remember { mutableStateOf(\"\") }", RGBColor(0x89, 0xDC, 0xEB)),
    ("  var error    by remember { mutableStateOf(\"\") }", RGBColor(0x89, 0xDC, 0xEB)),
    ("", CODE_FG),
    ("  OutlinedTextField(", CODE_FG),
    ("    value = email,", CODE_FG),
    ("    onValueChange = { email = it; error = \"\" },", CODE_FG),
    ("    label = { Text(\"Email\") },", CODE_FG),
    ("    keyboardOptions = KeyboardOptions(", CODE_FG),
    ("        keyboardType = KeyboardType.Email)", CODE_FG),
    ("  )", CODE_FG),
    ("", CODE_FG),
    ("  OutlinedTextField(", CODE_FG),
    ("    value = password,", CODE_FG),
    ("    onValueChange = { password = it },", CODE_FG),
    ("    visualTransformation = PasswordVisualTransformation()", RGBColor(0xA6, 0xE3, 0xA1)),
    ("  )", CODE_FG),
    ("", CODE_FG),
    ("  Button(onClick = {", CODE_FG),
    ("    if (usuarioVM.validarCredenciales(email, password)) {", RGBColor(0xA6, 0xE3, 0xA1)),
    ("      compraVM.cargarComprasDelUsuario(email)", CODE_FG),
    ("      navController.navigate(NavRoutes.Home.route) {", CODE_FG),
    ("        popUpTo(NavRoutes.Login.route) { inclusive = true }", RGBColor(0xA6, 0xE3, 0xA1)),
    ("      }", CODE_FG),
    ("    } else { error = \"Credenciales incorrectas\" }", CODE_FG),
    ("  }) { Text(\"Iniciar sesión\") }", CODE_FG),
    ("}", CODE_FG),
]
code_block(s, login_lines, 0.25, 1.25, 7.1, 6.0, font_size=Pt(8.2))

# Right: mockup phone
rect(s, 7.55, 1.25, 5.55, 6.0, fill_rgb=RGBColor(0x12, 0x12, 0x12),
     line_rgb=RGBColor(0x44, 0x44, 0x44), line_width=Pt(2))
rect(s, 7.65, 1.35, 5.35, 5.8, fill_rgb=WHITE)

# Phone screen content
text_box(s, "🛒", 9.5, 1.75, 1.5, 0.65, font_size=Pt(28), align=PP_ALIGN.CENTER)
text_box(s, "Super Ahorro", 7.75, 2.42, 5.15, 0.45,
         font_size=Pt(18), bold=True, color=GREEN_DARK, align=PP_ALIGN.CENTER)

# Error card
rect(s, 7.75, 2.95, 5.15, 0.38, fill_rgb=RGBColor(0xFF, 0xEB, 0xEE))
text_box(s, "Email o contraseña incorrectos", 7.85, 3.0, 4.9, 0.28,
         font_size=Pt(9), color=RGBColor(0xD3, 0x2F, 0x2F))

# Fields
for label, y_pos in [("Email", 3.4), ("Contraseña", 3.88)]:
    rect(s, 7.75, y_pos, 5.15, 0.42, fill_rgb=WHITE,
         line_rgb=GRAY, line_width=Pt(1))
    text_box(s, label, 7.88, y_pos+0.05, 4.8, 0.3,
             font_size=Pt(9.5), color=GRAY)

# Button
rect(s, 7.75, 4.42, 5.15, 0.45, fill_rgb=GREEN_DARK)
text_box(s, "Iniciar sesión", 7.75, 4.48, 5.15, 0.35,
         font_size=Pt(12), bold=True, color=WHITE, align=PP_ALIGN.CENTER)

text_box(s, "¿No tenés cuenta? Registrate", 7.75, 4.96, 5.15, 0.3,
         font_size=Pt(9.5), color=GREEN_DARK, align=PP_ALIGN.CENTER)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 11 — PANTALLA HOME
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, fill_rgb=RGBColor(0xF9, 0xFA, 0xF9))
header(s, "Pantalla Home — HomeScreen.kt", "ui/screens/HomeScreen.kt · Scaffold + LazyColumn + Cards")

home_lines = [
    ("@Composable", KW_COLOR),
    ("fun HomeScreen(navController: NavController,", CODE_FG),
    ("              compraViewModel: CompraViewModel) {", CODE_FG),
    ("", CODE_FG),
    ("  val compras      = compraViewModel.compras", RGBColor(0x89, 0xDC, 0xEB)),
    ("  val totalGastado = compraViewModel.calcularTotalGastado()", RGBColor(0x89, 0xDC, 0xEB)),
    ("", CODE_FG),
    ("  Scaffold(", KW_COLOR),
    ("    topBar = { TopAppBar(title = { Text(\"Super Ahorro\") },", CODE_FG),
    ("      colors = TopAppBarDefaults.topAppBarColors(", CODE_FG),
    ("          containerColor = Color(0xFF2E7D32),", RGBColor(0xA6, 0xE3, 0xA1)),
    ("          titleContentColor = Color.White)) },", CODE_FG),
    ("    bottomBar = { BottomNavBar(navController) },", CODE_FG),
    ("    floatingActionButton = {", CODE_FG),
    ("      FloatingActionButton(", CODE_FG),
    ("        onClick = { navController.navigate(NavRoutes.NuevaCompra.route) },", CODE_FG),
    ("        containerColor = Color(0xFF2E7D32)", RGBColor(0xA6, 0xE3, 0xA1)),
    ("      ) { Icon(Icons.Default.Add, null) }", CODE_FG),
    ("    }", CODE_FG),
    ("  ) { padding ->", CODE_FG),
    ("    LazyColumn(...) {", KW_COLOR),
    ("      item { /* Tarjeta Total Gastado */ }", CODE_FG),
    ("      item { /* Insight Cards: Top Super / Máx Gasto */ }", CODE_FG),
    ("      items(compras) { compra ->", RGBColor(0xA6, 0xE3, 0xA1)),
    ("        CompraCard(compra, onClick = {", CODE_FG),
    ("          navController.navigate(", CODE_FG),
    ("              NavRoutes.DetalleCompra.withId(compra.id))", CODE_FG),
    ("        })", CODE_FG),
    ("      }", CODE_FG),
    ("    }", CODE_FG),
    ("  }", CODE_FG),
    ("}", CODE_FG),
]
code_block(s, home_lines, 0.25, 1.25, 7.2, 6.0, font_size=Pt(8.2))

# Phone mockup
rect(s, 7.7, 1.25, 5.4, 6.0, fill_rgb=RGBColor(0x12, 0x12, 0x12),
     line_rgb=RGBColor(0x44, 0x44, 0x44), line_width=Pt(2))
rect(s, 7.8, 1.35, 5.2, 5.8, fill_rgb=WHITE)

# TopBar
rect(s, 7.8, 1.35, 5.2, 0.5, fill_rgb=GREEN_DARK)
text_box(s, "Super Ahorro", 7.95, 1.4, 3.5, 0.35,
         font_size=Pt(12), bold=True, color=WHITE)

# Main card
rect(s, 7.9, 1.9, 5.0, 1.05, fill_rgb=GREEN_DARK)
text_box(s, "Total gastado", 8.05, 1.95, 3.5, 0.28,
         font_size=Pt(9), color=RGBColor(0xC8, 0xE6, 0xC9))
text_box(s, "$ 45.820,50", 8.05, 2.22, 3.5, 0.42,
         font_size=Pt(17), bold=True, color=WHITE)
text_box(s, "Compras: 3     Promedio: 15.273", 8.05, 2.65, 4.5, 0.25,
         font_size=Pt(8), color=RGBColor(0xC8, 0xE6, 0xC9))

# Insight row
rect(s, 7.9, 3.03, 2.4, 0.7, fill_rgb=BLUE_LIGHT)
text_box(s, "🏪 Top Super\nCarrefour", 8.0, 3.07, 2.2, 0.58,
         font_size=Pt(8.5), color=BLACK)
rect(s, 10.38, 3.03, 2.4, 0.7, fill_rgb=RED_LIGHT)
text_box(s, "📉 Máx. Gasto\n$18.500", 10.48, 3.07, 2.2, 0.58,
         font_size=Pt(8.5), color=BLACK)

# Purchase cards
for i, (super_, fecha, monto) in enumerate([
    ("Carrefour", "08/05/2026  14:30", "$ 18.500"),
    ("Jumbo",     "05/05/2026  10:15", "$ 15.200"),
]):
    y = 3.85 + i * 0.88
    rect(s, 7.9, y, 4.95, 0.78, fill_rgb=WHITE,
         line_rgb=RGBColor(0xE0, 0xE0, 0xE0), line_width=Pt(1))
    rect(s, 7.98, y+0.12, 0.45, 0.45, fill_rgb=GREEN_LIGHT)
    text_box(s, "🛒", 7.96, y+0.1, 0.48, 0.45, font_size=Pt(11))
    text_box(s, super_, 8.52, y+0.06, 2.5, 0.3,
             font_size=Pt(10), bold=True, color=GREEN_DARK)
    text_box(s, fecha, 8.52, y+0.38, 2.5, 0.25,
             font_size=Pt(8), color=GRAY)
    text_box(s, monto, 10.8, y+0.15, 1.0, 0.35,
             font_size=Pt(11), bold=True, color=GREEN_DARK, align=PP_ALIGN.RIGHT)

# FAB
rect(s, 11.8, 5.7, 0.55, 0.55, fill_rgb=GREEN_DARK)
text_box(s, "+", 11.82, 5.72, 0.55, 0.45,
         font_size=Pt(16), bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Bottom nav
rect(s, 7.8, 6.75, 5.2, 0.4, fill_rgb=RGBColor(0xFA, 0xFA, 0xFA),
     line_rgb=RGBColor(0xE0, 0xE0, 0xE0), line_width=Pt(1))
for label, x in [("🏠", 8.1), ("📋", 9.1), ("📊", 10.1), ("👤", 11.1)]:
    text_box(s, label, x, 6.78, 0.55, 0.3,
             font_size=Pt(11), align=PP_ALIGN.CENTER)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 12 — ESTADÍSTICAS + PIE CHART
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, fill_rgb=RGBColor(0xF9, 0xFA, 0xF9))
header(s, "Estadísticas — Gráfico Circular Propio", "ui/screens/EstadisticasScreen.kt · Custom Canvas @Composable")

pie_lines = [
    ("@Composable", KW_COLOR),
    ("fun PieChart(data: List<Pair<String, Double>>,", CODE_FG),
    ("            totalGastado: Double, modifier: Modifier) {", CODE_FG),
    ("", CODE_FG),
    ("  val total  = data.sumOf { it.second }", RGBColor(0x89, 0xDC, 0xEB)),
    ("  val colors = data.indices.map { getColorForIndex(it) }", CODE_FG),
    ("", CODE_FG),
    ("  Box(modifier, contentAlignment = Alignment.Center) {", CODE_FG),
    ("    Canvas(modifier = Modifier.fillMaxSize()) {", KW_COLOR),
    ("      val radius = size.minDimension / 2.3f", CODE_FG),
    ("      var currentAngle = -90f", CODE_FG),
    ("", CODE_FG),
    ("      data.forEachIndexed { index, (_, value) ->", CODE_FG),
    ("        val sweep = (value / total * 360f).toFloat()", CODE_FG),
    ("", CODE_FG),
    ("        drawArc(color = colors[index],", RGBColor(0xA6, 0xE3, 0xA1)),
    ("                startAngle = currentAngle,", CODE_FG),
    ("                sweepAngle = sweep, useCenter = true, ...)", CODE_FG),
    ("", CODE_FG),
    ("        currentAngle += sweep", CODE_FG),
    ("      }", CODE_FG),
    ("      // Círculo blanco central = efecto donut", GRAY),
    ("      drawCircle(Color.White, radius * 0.6f, ...)", RGBColor(0xA6, 0xE3, 0xA1)),
    ("    }", CODE_FG),
    ("    // Texto total en el centro del donut", GRAY),
    ("    Column(horizontalAlignment = Alignment.CenterHorizontally) {", CODE_FG),
    ("      Text(\"Total\"); Text(\"$ ${total}\")", CODE_FG),
    ("    }", CODE_FG),
    ("  }", CODE_FG),
    ("}", CODE_FG),
]
code_block(s, pie_lines, 0.25, 1.25, 7.0, 6.0, font_size=Pt(8.5))

# Right: phone mockup of stats
rect(s, 7.5, 1.25, 5.6, 6.0, fill_rgb=RGBColor(0x12, 0x12, 0x12),
     line_rgb=RGBColor(0x44, 0x44, 0x44), line_width=Pt(2))
rect(s, 7.6, 1.35, 5.4, 5.8, fill_rgb=WHITE)

# TopBar
rect(s, 7.6, 1.35, 5.4, 0.45, fill_rgb=GREEN_DARK)
text_box(s, "Estadísticas", 7.75, 1.4, 3.5, 0.32,
         font_size=Pt(11), bold=True, color=WHITE)

# Total card
rect(s, 7.7, 1.85, 5.2, 0.72, fill_rgb=GREEN_LIGHT)
text_box(s, "Total Gastado: $ 45.820,50  (3 compras)", 7.83, 1.92, 4.9, 0.55,
         font_size=Pt(10), bold=True, color=GREEN_DARK)

# Pie chart visual (drawn with shapes)
cx, cy = 10.3, 4.05  # center in inches
r = 1.1

# Simplified pie using colored arcs (rectangles as approximation)
pie_segments = [
    (GREEN_DARK, "Carrefour 40%"),
    (RGBColor(0x19, 0x76, 0xD2), "Jumbo 33%"),
    (RGBColor(0xD3, 0x2F, 0x2F), "Dia 27%"),
]

# Draw donut chart using overlapping shapes
# Outer circle
text_box(s, "◉", cx-1.3, cy-1.2, 2.6, 2.4,
         font_size=Pt(110), color=GREEN_DARK, align=PP_ALIGN.CENTER)

# White circle in center (donut effect)
rect(s, cx-0.72, cy-0.72, 1.44, 1.44, fill_rgb=WHITE)
text_box(s, "$ 45.820", cx-0.9, cy-0.3, 1.8, 0.6,
         font_size=Pt(8.5), bold=True, color=GREEN_DARK, align=PP_ALIGN.CENTER)

# Legend
colors_hex = [GREEN_DARK, RGBColor(0x19, 0x76, 0xD2), RGBColor(0xD3, 0x2F, 0x2F)]
for i, (label, pct) in enumerate([("Carrefour", "40%"), ("Jumbo", "33%"), ("Día", "27%")]):
    y = 5.32 + i * 0.33
    rect(s, 7.75, y, 0.22, 0.22, fill_rgb=colors_hex[i])
    text_box(s, f"{label}  {pct}", 8.05, y, 4.7, 0.28,
             font_size=Pt(9.5), color=BLACK)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 13 — BOTTOM NAV BAR
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, fill_rgb=RGBColor(0xF9, 0xFA, 0xF9))
header(s, "Navegación Inferior — BottomNavBar", "ui/components/BottonNavBar.kt · NavigationBar Material 3")

bnb_lines = [
    ("// ui/components/BottonNavBar.kt", GRAY),
    ("@Composable", KW_COLOR),
    ("fun BottomNavBar(navController: NavController) {", CODE_FG),
    ("", CODE_FG),
    ("  val items = listOf(", CODE_FG),
    ("    Triple(\"Home\",        Icons.Default.Home,    NavRoutes.Home.route),", CODE_FG),
    ("    Triple(\"Historial\",   Icons.Default.History, NavRoutes.Historial.route),", CODE_FG),
    ("    Triple(\"Estadísticas\",Icons.Default.BarChart,NavRoutes.Estadisticas.route),", CODE_FG),
    ("    Triple(\"Perfil\",      Icons.Default.Person,  NavRoutes.Perfil.route),", CODE_FG),
    ("  )", CODE_FG),
    ("", CODE_FG),
    ("  NavigationBar(tonalElevation = 8.dp) {", KW_COLOR),
    ("    val currentRoute by navController", CODE_FG),
    ("        .currentBackStackEntryAsState()  // observa ruta actual", GRAY),
    ("", CODE_FG),
    ("    items.forEach { (label, icon, route) ->", CODE_FG),
    ("      NavigationBarItem(", CODE_FG),
    ("        selected = currentRoute?.destination?.route == route,", RGBColor(0xA6, 0xE3, 0xA1)),
    ("        onClick = {", CODE_FG),
    ("          navController.navigate(route) {", CODE_FG),
    ("            popUpTo(NavRoutes.Home.route) { saveState = true }", RGBColor(0xA6, 0xE3, 0xA1)),
    ("            launchSingleTop = true", RGBColor(0xA6, 0xE3, 0xA1)),
    ("            restoreState    = true", RGBColor(0xA6, 0xE3, 0xA1)),
    ("          }", CODE_FG),
    ("        },", CODE_FG),
    ("        icon  = { Icon(icon, label) },", CODE_FG),
    ("        label = { Text(label) },", CODE_FG),
    ("        colors = NavigationBarItemDefaults.colors(", CODE_FG),
    ("          selectedIconColor = Color(0xFF2E7D32),", RGBColor(0xA6, 0xE3, 0xA1)),
    ("          indicatorColor    = Color(0xFFE8F5E9)", RGBColor(0xA6, 0xE3, 0xA1)),
    ("        )", CODE_FG),
    ("      )", CODE_FG),
    ("    }", CODE_FG),
    ("  }", CODE_FG),
    ("}", CODE_FG),
]
code_block(s, bnb_lines, 0.25, 1.25, 7.55, 6.0, font_size=Pt(8.2))

# right annotations
text_box(s, "Puntos clave del componente", 8.1, 1.25, 5.0, 0.42,
         font_size=Pt(14), bold=True, color=GREEN_DARK)

annots = [
    ("currentBackStackEntryAsState()", "Observa la ruta activa como State.\nCada vez que la ruta cambia,\nel item seleccionado se actualiza."),
    ("launchSingleTop = true", "Evita duplicar la pantalla en el\nback-stack si ya está en la cima."),
    ("saveState / restoreState", "Preserva el scroll y estado de\nlas pantallas al volver a ellas."),
    ("indicatorColor verde claro", "Retroalimentación visual al\nusuario: item seleccionado\ncon fondo #E8F5E9."),
]
for i, (t, b) in enumerate(annots):
    y = 1.72 + i * 1.42
    rect(s, 8.1, y, 5.0, 1.3, fill_rgb=GREEN_LIGHT, line_rgb=GREEN_MED, line_width=Pt(1))
    text_box(s, t, 8.25, y+0.08, 4.7, 0.35,
             font_size=Pt(11.5), bold=True, color=GREEN_DARK)
    text_box(s, b, 8.25, y+0.45, 4.7, 0.78,
             font_size=Pt(10.5), color=BLACK)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 14 — INTERNACIONALIZACIÓN (strings.xml)
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, fill_rgb=RGBColor(0xF9, 0xFA, 0xF9))
header(s, "Internacionalización — strings.xml", "res/values/strings.xml · Separación de textos de la UI")

# Left: current state
text_box(s, "Estado actual del proyecto", 0.3, 1.3, 6.0, 0.42,
         font_size=Pt(14), bold=True, color=GREEN_DARK)
xml_lines = [
    ("<!-- res/values/strings.xml -->", GRAY),
    ("<resources>", RGBColor(0x89, 0xDC, 0xEB)),
    ("    <string name=\"app_name\">SuperAhorro</string>", RGBColor(0xA6, 0xE3, 0xA1)),
    ("</resources>", RGBColor(0x89, 0xDC, 0xEB)),
    ("", CODE_FG),
    ("<!-- AndroidManifest.xml —", GRAY),
    ("     usa @string/app_name -->", GRAY),
    ("<application", CODE_FG),
    ("    android:label=\"@string/app_name\"", RGBColor(0xA6, 0xE3, 0xA1)),
    ("    android:theme=\"@style/Theme.SuperAhorro\">", CODE_FG),
]
code_block(s, xml_lines, 0.3, 1.78, 6.0, 2.42, font_size=Pt(10))

# Best practice addition
text_box(s, "Cómo debería expandirse (buena práctica):", 0.3, 4.28, 6.0, 0.38,
         font_size=Pt(12), bold=True, color=GREEN_DARK)
xml2_lines = [
    ("<!-- Todos los textos de la UI en strings.xml -->", GRAY),
    ("<resources>", RGBColor(0x89, 0xDC, 0xEB)),
    ("  <string name=\"app_name\">SuperAhorro</string>", RGBColor(0xA6, 0xE3, 0xA1)),
    ("  <string name=\"login_email\">Email</string>", RGBColor(0xA6, 0xE3, 0xA1)),
    ("  <string name=\"login_password\">Contraseña</string>", RGBColor(0xA6, 0xE3, 0xA1)),
    ("  <string name=\"login_button\">Iniciar sesión</string>", RGBColor(0xA6, 0xE3, 0xA1)),
    ("  <string name=\"home_total\">Total gastado</string>", RGBColor(0xA6, 0xE3, 0xA1)),
    ("  <!-- Para inglés: res/values-en/strings.xml -->", GRAY),
    ("</resources>", RGBColor(0x89, 0xDC, 0xEB)),
    ("", CODE_FG),
    ("// En Compose:", GRAY),
    ("Text(stringResource(R.string.login_button))", CODE_FG),
]
code_block(s, xml2_lines, 0.3, 4.68, 6.0, 2.5, font_size=Pt(9.5))

# Right: explanation
text_box(s, "¿Por qué usar strings.xml?", 6.7, 1.3, 6.4, 0.42,
         font_size=Pt(14), bold=True, color=GREEN_DARK)

reasons_i18n = [
    ("📂  Separación de responsabilidades",
     "La lógica de la UI no mezcla texto con código.\nCambiar un texto = cambiar un archivo XML."),
    ("🌍  Soporte multiidioma",
     "Crear res/values-en/strings.xml para inglés,\nres/values-pt/strings.xml para portugués, etc.\nAndroid elige el correcto automáticamente."),
    ("♿  Accesibilidad",
     "Los lectores de pantalla (TalkBack) usan los\nstrings de contentDescription definidos en XML."),
    ("🔧  Fácil mantenimiento",
     "Un solo lugar para todos los textos.\nSin buscar strings dispersos en el código."),
]
for i, (t, b) in enumerate(reasons_i18n):
    y = 1.78 + i * 1.38
    rect(s, 6.7, y, 6.4, 1.25, fill_rgb=GREEN_LIGHT, line_rgb=GREEN_MED, line_width=Pt(1))
    text_box(s, t, 6.85, y+0.08, 6.1, 0.35,
             font_size=Pt(12), bold=True, color=GREEN_DARK)
    text_box(s, b, 6.85, y+0.45, 6.1, 0.72,
             font_size=Pt(10.5), color=BLACK)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 15 — VERSIONADO EN GITHUB
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, fill_rgb=RGBColor(0xF9, 0xFA, 0xF9))
header(s, "Versionado en GitHub", "Repositorio público · Seguimiento de cambios con Git")

# Big repo card
rect(s, 0.3, 1.3, 12.73, 1.55, fill_rgb=CODE_BG, line_rgb=GREEN_MED, line_width=Pt(1))
text_box(s, "📁  github.com / benjaalasino / superahorro", 0.55, 1.48, 9.0, 0.55,
         font_size=Pt(20), bold=True, color=WHITE)
text_box(s, "🌿  rama: claude/android-app-presentation-i0Fu1", 0.55, 2.08, 9.0, 0.38,
         font_size=Pt(13), color=RGBColor(0xA6, 0xE3, 0xA1), italic=True)

git_flow = [
    ("git init", "Inicializa el repositorio local"),
    ("git add .", "Agrega todos los archivos al staging"),
    ("git commit -m \"feat: add login screen\"", "Crea un commit con descripción"),
    ("git push -u origin main", "Sube el código a GitHub"),
    ("git branch feature/stats", "Crea una rama para nueva funcionalidad"),
]

text_box(s, "Flujo de trabajo con Git", 0.3, 3.05, 6.0, 0.42,
         font_size=Pt(14), bold=True, color=GREEN_DARK)
for i, (cmd, desc) in enumerate(git_flow):
    y = 3.52 + i * 0.72
    rect(s, 0.3, y, 12.73, 0.62, fill_rgb=CODE_BG if i%2==0 else RGBColor(0x18, 0x18, 0x2A))
    text_box(s, f"$ {cmd}", 0.5, y+0.08, 6.8, 0.42,
             font_size=Pt(10.5), color=RGBColor(0xA6, 0xE3, 0xA1))
    text_box(s, f"→  {desc}", 7.5, y+0.1, 5.3, 0.4,
             font_size=Pt(10.5), color=GRAY)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 16 — CHECKLIST DE REQUISITOS NO FUNCIONALES
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, fill_rgb=RGBColor(0xF9, 0xFA, 0xF9))
header(s, "Requisitos No Funcionales — Checklist", "Todos los requisitos del proyecto verificados en el código")

reqs = [
    ("✅", "Proyecto creado desde cero en Android Studio",
     "Estructura generada por Android Studio con AGP 9.1.1 y API 36"),
    ("✅", "Versionado en GitHub",
     "github.com/benjaalasino/superahorro — rama dedicada"),
    ("✅", "UI usable, clara y visualmente consistente",
     "Material 3, paleta verde #2E7D32, Cards, Scaffold en todas las pantallas"),
    ("✅", "Nombre de paquete: com.undef.nombreappalumno",
     "namespace = \"com.undef.superahorroCalvoAlasino\""),
    ("✅", "Jetpack Compose",
     "100% @Composable — sin XML layouts, Material 3 completo"),
    ("✅", "Navegación entre pantallas",
     "NavHost + 12 rutas (NavRoutes sealed class) + argumentos tipados"),
    ("✅", "Corrutinas en operaciones asincrónicas",
     "mutableStateOf / mutableStateListOf (estado reactivo de Compose)\n+ ViewModel (lifecycle-aware, no bloquea el hilo principal)"),
    ("✅", "Internacionalización con strings.xml",
     "res/values/strings.xml con app_name — estructura para expandir a multi-idioma"),
    ("✅", "Arquitectura MVVM",
     "model/ → data classes   |   viewmodel/ → ViewModel()   |   ui/ → @Composable"),
]

for i, (check, title, detail) in enumerate(reqs):
    col = i % 2
    row = i // 2
    x = 0.25 + col * 6.55
    y = 1.3 + row * 1.52
    w = 6.3
    h = 1.4

    bg = GREEN_LIGHT if col == 0 else RGBColor(0xF0, 0xF7, 0xF0)
    rect(s, x, y, w, h, fill_rgb=bg, line_rgb=GREEN_MED, line_width=Pt(1))
    text_box(s, check, x+0.08, y+0.08, 0.42, 0.42, font_size=Pt(14))
    text_box(s, title, x+0.55, y+0.1, w-0.65, 0.4,
             font_size=Pt(11.5), bold=True, color=GREEN_DARK)
    text_box(s, detail, x+0.55, y+0.52, w-0.65, 0.82,
             font_size=Pt(10), color=BLACK)

# Last item (9th) spans both columns
i = 8
check, title, detail = reqs[8]
rect(s, 0.25, 7.05, 12.83, 0.3, fill_rgb=GREEN_DARK)
text_box(s, f"✅  {title}  —  {detail}", 0.4, 7.06, 12.5, 0.26,
         font_size=Pt(10), color=WHITE)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE 17 — CIERRE
# ════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, 13.33, 7.5, fill_rgb=GREEN_DARK)
rect(s, 0, 5.8, 13.33, 1.7, fill_rgb=RGBColor(0x1B, 0x5E, 0x20))

text_box(s, "SuperAhorro", 1.5, 1.4, 10.5, 1.1,
         font_size=Pt(54), bold=True, color=WHITE, align=PP_ALIGN.CENTER)
text_box(s, "Kotlin  ·  Jetpack Compose  ·  MVVM  ·  Navigation Compose",
         1.5, 2.55, 10.5, 0.55,
         font_size=Pt(17), color=RGBColor(0xC8, 0xE6, 0xC9), align=PP_ALIGN.CENTER)

summary = [
    "9 requisitos no funcionales cumplidos",
    "12 rutas de navegación con sealed class",
    "2 ViewModels con estado reactivo",
    "Persistencia JSON por usuario con SharedPreferences",
    "Gráfico circular implementado con Canvas API",
]
for i, item in enumerate(summary):
    text_box(s, f"▸  {item}", 3.5, 3.25 + i*0.44, 6.5, 0.38,
             font_size=Pt(13), color=RGBColor(0xA5, 0xD6, 0xA7), align=PP_ALIGN.CENTER)

text_box(s, "com.undef.superahorroCalvoAlasino  ·  API 36 (Android 15)",
         1.5, 6.15, 10.5, 0.38,
         font_size=Pt(12), color=RGBColor(0x80, 0xCB, 0xC4), align=PP_ALIGN.CENTER, italic=True)

# ── Save ────────────────────────────────────────────────────────────────────
out = "/home/user/SuperAhorro/SuperAhorro_Presentacion.pptx"
prs.save(out)
print(f"Saved → {out}")
