"""Page Object de ValidaCarga: todo lo que las pruebas necesitan para usar la página.
Las pruebas dicen QUÉ se espera; esta clase sabe CÓMO se hace en la interfaz."""
import pathlib
from playwright.sync_api import Page, expect


class ValidaCarga:
    def __init__(self, page: Page):
        self.page = page
        self.t = page.get_by_test_id

    # ------------------------------------------------------------ navegación
    def abrir(self, url: str):
        self.page.goto(url)
        expect(self.t('resumen-total')).to_be_visible()  # abre con la plantilla demo cargada
        return self

    def limpiar(self):
        self.t('btn-limpiar').click()
        if self.modo_actual() == 'plantilla':
            expect(self.t('zona-plantilla')).to_be_visible()

    def modo(self, nombre: str):
        self.t(f'modo-{nombre}').click()

    def modo_actual(self) -> str:
        return 'plantilla' if self.t('modo-plantilla').get_attribute('aria-pressed') == 'true' else 'ejemplos'

    def perfil(self, perfil_id: str):
        self.t(f'perfil-{perfil_id}').click()

    # ------------------------------------------------------------ archivos
    def subir_plantilla(self, ruta: pathlib.Path):
        self.t('input-plantilla').set_input_files(str(ruta))

    def subir_archivo(self, ruta: pathlib.Path, esperar=True):
        self.t('input-archivo').set_input_files(str(ruta))
        if esperar:
            expect(self.t('nombre-archivo')).to_have_text(ruta.name)

    def probar_ejemplo(self):
        self.t('btn-muestra').click()

    def descargar(self, testid: str, carpeta: pathlib.Path) -> pathlib.Path:
        with self.page.expect_download() as info:
            self.t(testid).click()
        destino = carpeta / info.value.suggested_filename
        info.value.save_as(destino)
        return destino

    # ------------------------------------------------------------ lectura de resultados
    def resumen(self) -> dict:
        return {k: int(self.t(f'resumen-{k}').inner_text()) for k in ('total', 'correcto', 'advertencia', 'error')}

    def estado_fila(self, n: int) -> str:
        return self.t(f'fila-{n}').get_attribute('data-estado')

    def estados(self) -> dict:
        pares = self.page.eval_on_selector_all('tbody.grupo', 'els => els.map(e => [e.dataset.fila, e.dataset.estado])')
        return {int(f): e for f, e in pares}

    def texto_fila(self, n: int) -> str:
        return self.t(f'fila-{n}').inner_text()

    def mensaje(self) -> str:
        return self.t('mensaje').inner_text()

    def revision_plantilla(self) -> str:
        return self.t('revision-plantilla').inner_text()

    def reglas(self) -> list:
        """[(columna, texto de los chips, títulos de los chips, obligatoria)] del panel Reglas detectadas."""
        return self.page.eval_on_selector_all('[data-testid^="regla-"]', '''els => els.map(e => [
            e.querySelector('.regla-col span').textContent,
            [...e.querySelectorAll('.chip')].map(c => c.textContent).join(' | '),
            [...e.querySelectorAll('.chip')].map(c => c.title).join(' | '),
            e.querySelector('input[type=checkbox]').checked])''')

    def regla(self, columna: str):
        return next(r for r in self.reglas() if r[0] == columna)

    def opcion_elegida(self, testid: str) -> str:
        return self.t(testid).evaluate('e => e.options[e.selectedIndex].text')

    def incidencias_en_pantalla(self) -> int:
        return self.page.locator('span.sev').count()
