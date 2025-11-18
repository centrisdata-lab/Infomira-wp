"""
GESTOR DE COMUNIDADES DE WHATSAPP
===================================
Agrega y elimina personas de comunidades de WhatsApp de forma automatizada
"""

import os
import sys
import time
import random
import re
import pandas as pd
from datetime import datetime

def instalar_dependencias():
    """Instalar dependencias necesarias"""
    paquetes_requeridos = [
        'selenium==4.15.2',
        'webdriver-manager==4.0.1',
        'pandas',
        'openpyxl'
    ]

    print("🔍 Verificando dependencias...")

    for paquete in paquetes_requeridos:
        nombre_paquete = paquete.split('==')[0]
        try:
            __import__(nombre_paquete.replace('-', '_'))
            print(f"✅ {nombre_paquete} ya está instalado")
        except ImportError:
            print(f"📦 Instalando {paquete}...")
            import subprocess
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', paquete])

# Instalar dependencias
instalar_dependencias()

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager


class GestorComunidadesWhatsApp:
    def __init__(self):
        self.driver = None
        self.wait = None
        self.usar_cache = False
        self.tiempo_min_contacto = 5
        self.tiempo_max_contacto = 10
        self.tiempo_entre_procesos = 15
        self.session_path = os.path.join(os.getcwd(), "whatsapp_session")
        self.cantidad_procesar = None  # Cantidad de registros a procesar

    def configurar_parametros(self):
        """Configurar parámetros de tiempo y sesión"""
        print("\n" + "="*60)
        print("⚙️ CONFIGURACIÓN DE PARÁMETROS")
        print("="*60)

        # Verificar si existe sesión guardada
        tiene_cache = os.path.exists(self.session_path) and len(os.listdir(self.session_path)) > 0

        if tiene_cache:
            print("\n✅ Se encontró una sesión anterior guardada")
            print("¿Qué deseas hacer?")
            print("  1. Usar sesión guardada (no requiere escanear QR)")
            print("  2. Escanear nuevo QR (nueva sesión)")

            opcion_sesion = input("\nElige una opción (1/2): ").strip()
            self.usar_cache = (opcion_sesion == "1")

            if not self.usar_cache:
                print("🔄 Se escaneará un nuevo QR y se guardará la sesión")
        else:
            print("\n📱 Es la primera vez que usas esta herramienta")
            print("✅ Deberás escanear el código QR de WhatsApp")
            self.usar_cache = False

        # Configurar tiempos
        print("\n" + "="*60)
        print("⏱️ CONFIGURACIÓN DE TIEMPOS")
        print("="*60)
        print("\n📊 Tiempos entre contactos (agregar/eliminar):")
        print("   Estos son los tiempos de espera entre procesar cada contacto")

        try:
            self.tiempo_min_contacto = int(input("   ⏳ Tiempo MÍNIMO entre contactos (segundos, recomendado 5): ") or "5")
            self.tiempo_max_contacto = int(input("   ⏳ Tiempo MÁXIMO entre contactos (segundos, recomendado 10): ") or "10")
        except:
            print("   ⚠️ Valores inválidos, usando valores por defecto (5-10 segundos)")
            self.tiempo_min_contacto = 5
            self.tiempo_max_contacto = 10

        print(f"\n✅ Tiempo entre contactos: {self.tiempo_min_contacto}-{self.tiempo_max_contacto} segundos")
        print(f"✅ Tiempo entre procesos (agregar→eliminar): {self.tiempo_entre_procesos} segundos")

        # Configurar cantidad de registros a procesar
        print("\n" + "="*60)
        print("📊 CANTIDAD DE REGISTROS A PROCESAR")
        print("="*60)
        print("¿Cuántos registros quieres procesar?")
        print("  1. Solo 3 (prueba)")
        print("  2. Todos")
        print("  3. Cantidad personalizada")

        opcion_cantidad = input("\nElige una opción (1/2/3): ").strip()

        if opcion_cantidad == "1":
            self.cantidad_procesar = 3
            print("✅ Se procesarán 3 registros (modo prueba)")
        elif opcion_cantidad == "3":
            try:
                cantidad = int(input("¿Cuántos registros?: "))
                self.cantidad_procesar = cantidad
                print(f"✅ Se procesarán {cantidad} registros")
            except:
                self.cantidad_procesar = 3
                print("⚠️ Valor inválido, se procesarán 3 registros")
        else:
            self.cantidad_procesar = None  # None significa todos
            print("✅ Se procesarán TODOS los registros")

    def limpiar_texto_para_selenium(self, texto):
        """Limpiar emojis y caracteres especiales que ChromeDriver no puede manejar"""
        try:
            # Remover emojis y caracteres fuera del BMP (Basic Multilingual Plane)
            # Mantener solo caracteres ASCII extendido y algunos Unicode básicos
            texto_limpio = ''.join(char for char in texto if ord(char) < 0x10000)

            # Si quedó vacío o muy corto, retornar el original
            if len(texto_limpio.strip()) < 3:
                # Intentar método alternativo: remover solo emojis comunes
                emoji_pattern = re.compile("["
                    u"\U0001F600-\U0001F64F"  # emoticones
                    u"\U0001F300-\U0001F5FF"  # símbolos & pictogramas
                    u"\U0001F680-\U0001F6FF"  # transporte & símbolos de mapa
                    u"\U0001F1E0-\U0001F1FF"  # banderas (iOS)
                    u"\U00002702-\U000027B0"
                    u"\U000024C2-\U0001F251"
                    u"\U0001f926-\U0001f937"
                    u"\U00010000-\U0010ffff"
                    u"\u2640-\u2642"
                    u"\u2600-\u2B55"
                    u"\u200d"
                    u"\u23cf"
                    u"\u23e9"
                    u"\u231a"
                    u"\ufe0f"  # dingbats
                    u"\u3030"
                    "]+", flags=re.UNICODE)
                texto_limpio = emoji_pattern.sub(r'', texto)

            return texto_limpio.strip()
        except:
            # Si falla, retornar texto sin modificar
            return texto

    def configurar_navegador(self):
        """Configurar navegador Chrome con WhatsApp Web"""
        try:
            print("\n🌐 Configurando navegador...")

            options = Options()

            # Configurar perfil de usuario para mantener sesión
            options.add_argument(f"--user-data-dir={self.session_path}")
            options.add_argument("--profile-directory=Default")

            # Configuración para parecer más humano
            options.add_argument("--start-maximized")
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option('useAutomationExtension', False)

            # Preferencias adicionales
            prefs = {
                "profile.default_content_setting_values.notifications": 2,  # Bloquear notificaciones
                "profile.default_content_setting_values.media_stream": 1,   # Permitir micrófono/cámara si es necesario
            }
            options.add_experimental_option("prefs", prefs)

            # Configurar servicio
            service = Service(ChromeDriverManager().install())

            # Inicializar driver
            self.driver = webdriver.Chrome(service=service, options=options)
            self.wait = WebDriverWait(self.driver, 30)

            # Script anti-detección
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

            print("✅ Navegador configurado correctamente")
            return True

        except Exception as e:
            print(f"❌ Error configurando navegador: {e}")
            return False

    def iniciar_whatsapp(self):
        """Iniciar WhatsApp Web"""
        try:
            print("\n📱 Abriendo WhatsApp Web...")
            self.driver.get("https://web.whatsapp.com")

            if not self.usar_cache:
                print("\n" + "="*60)
                print("📱 ESCANEA EL CÓDIGO QR")
                print("="*60)
                print("1. Abre WhatsApp en tu teléfono")
                print("2. Ve a Configuración > Dispositivos vinculados")
                print("3. Escanea el código QR que aparece en la ventana del navegador")
                print("4. La sesión se guardará para usos futuros")
                print("="*60)
            else:
                print("✅ Usando sesión guardada, no necesitas escanear QR")

            # Esperar a que cargue WhatsApp Web
            print("⏳ Esperando que WhatsApp Web cargue (puede tomar hasta 90 segundos)...")
            try:
                # Aumentar timeout específicamente para esta operación
                wait_largo = WebDriverWait(self.driver, 90)

                # Esperar por el buscador de chats (indica que está logueado)
                wait_largo.until(EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='3']")))
                print("✅ WhatsApp Web cargado exitosamente")
                time.sleep(3)
                return True
            except Exception as e:
                print("❌ No se pudo cargar WhatsApp Web")
                print(f"   Error: {e}")
                print("   Verifica que hayas escaneado el QR correctamente")
                return False

        except Exception as e:
            print(f"❌ Error iniciando WhatsApp: {e}")
            return False

    def esperar_aleatorio(self, min_seg, max_seg):
        """Esperar un tiempo aleatorio para simular comportamiento humano"""
        tiempo = random.uniform(min_seg, max_seg)
        print(f"⏳ Esperando {tiempo:.1f} segundos...")
        time.sleep(tiempo)

    def _cerrar_ventanas_modales(self):
        """Cerrar todas las ventanas modales y volver a la vista principal de chat"""
        try:
            print("  🔄 Cerrando ventanas modales...")

            # Presionar ESC varias veces para asegurar que todo se cierra
            for i in range(3):
                try:
                    ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
                    time.sleep(0.5)
                except:
                    pass

            # Verificar si hay botones de cerrar (X) visibles
            try:
                botones_cerrar = self.driver.find_elements(
                    By.XPATH,
                    "//button[@aria-label='Cerrar' or @aria-label='Close' or contains(@aria-label, 'cerrar')]"
                )
                for boton in botones_cerrar:
                    try:
                        if boton.is_displayed():
                            boton.click()
                            time.sleep(0.5)
                    except:
                        pass
            except:
                pass

            print("  ✓ Ventanas modales cerradas")
            time.sleep(1)
            return True

        except Exception as e:
            print(f"  ⚠️ Error cerrando ventanas: {e}")
            return False

    def buscar_comunidad(self, nombre_comunidad):
        """Buscar y abrir una comunidad"""
        try:
            print(f"\n🔍 Buscando comunidad: {nombre_comunidad}")

            # Limpiar el nombre de emojis para evitar errores de ChromeDriver
            nombre_limpio = self.limpiar_texto_para_selenium(nombre_comunidad)
            if nombre_limpio != nombre_comunidad:
                print(f"   ⚠️ Nombre con emojis detectado, usando versión limpia: {nombre_limpio}")

            # Hacer clic en el buscador
            buscador = self.wait.until(EC.presence_of_element_located(
                (By.XPATH, "//div[@contenteditable='true'][@data-tab='3']")
            ))
            buscador.click()
            time.sleep(1)

            # Limpiar completamente el buscador (importante para nuevas búsquedas)
            # Método 1: Usar clear()
            buscador.clear()
            time.sleep(0.5)

            # Método 2: Seleccionar todo y borrar (más confiable en WhatsApp Web)
            buscador.send_keys(Keys.CONTROL + "a")
            time.sleep(0.3)
            buscador.send_keys(Keys.DELETE)
            time.sleep(0.5)

            print(f"   ✓ Buscador limpiado")

            # Escribir el nombre de la comunidad (versión limpia)
            buscador.send_keys(nombre_limpio)
            print(f"   ✓ Buscando: {nombre_limpio}")
            self.esperar_aleatorio(2, 3)

            # Buscar el resultado y hacer clic
            try:
                resultado = None

                # Intento 1: Buscar por el span con el título que contiene el nombre
                try:
                    span_resultado = self.wait.until(EC.presence_of_element_located(
                        (By.XPATH, f"//span[contains(@title, '{nombre_limpio}')]")
                    ))
                    # Buscar el div padre clickeable (el contenedor del chat)
                    resultado = span_resultado.find_element(By.XPATH, "./ancestor::div[@role='listitem' or @role='row'][1]")
                    print(f"   ✓ Resultado encontrado (método 1: por título)")
                except Exception as e1:
                    print(f"   ℹ️ Intento 1 falló: {e1}")
                    pass

                # Intento 2: Buscar el primer resultado en la lista
                if not resultado:
                    try:
                        resultado = self.wait.until(EC.presence_of_element_located(
                            (By.XPATH, "//div[@id='pane-side']//div[@role='listitem'][1] | //div[@id='pane-side']//div[@role='row'][1]")
                        ))
                        print(f"   ✓ Resultado encontrado (método 2: primer item)")
                    except Exception as e2:
                        print(f"   ℹ️ Intento 2 falló: {e2}")
                        pass

                # Intento 3: Presionar Enter en el buscador
                if not resultado:
                    try:
                        print(f"   ℹ️ Intentando con Enter...")
                        buscador.send_keys(Keys.ENTER)
                        time.sleep(2)
                        print(f"   ✓ Enter presionado")
                        # Verificar si se abrió el chat
                        try:
                            self.driver.find_element(By.XPATH, "//header[@data-testid='conversation-header']")
                            print(f"✅ Comunidad '{nombre_comunidad}' abierta (método Enter)")
                            self.esperar_aleatorio(2, 3)
                            return True
                        except:
                            print(f"   ⚠️ Enter no abrió el chat")
                            pass
                    except Exception as e3:
                        print(f"   ℹ️ Intento 3 falló: {e3}")
                        pass

                if resultado:
                    # Intentar hacer clic - Múltiples métodos
                    clic_exitoso = False

                    # Método 1: Doble clic (más confiable en WhatsApp)
                    try:
                        ActionChains(self.driver).double_click(resultado).perform()
                        print(f"   ✓ Doble clic en resultado")
                        time.sleep(3)

                        # Verificar si se abrió - buscar header o contenido de chat
                        try:
                            # Verificar si existe header de conversación
                            self.driver.find_element(By.XPATH, "//header[@data-testid='conversation-header']")
                            clic_exitoso = True
                        except:
                            # Si no, verificar si hay contenido de chat visible (área de mensajes)
                            try:
                                self.driver.find_element(By.XPATH, "//div[@data-testid='conversation-panel-body'] | //div[contains(@class, 'copyable-area')]")
                                clic_exitoso = True
                            except:
                                pass
                    except Exception as e:
                        print(f"   ℹ️ Doble clic falló: {e}")

                    # Método 2: Clic simple si el doble clic no funcionó
                    if not clic_exitoso:
                        try:
                            resultado.click()
                            print(f"   ✓ Clic simple en resultado")
                            time.sleep(3)

                            # Verificar si se abrió
                            try:
                                self.driver.find_element(By.XPATH, "//header[@data-testid='conversation-header']")
                                clic_exitoso = True
                            except:
                                try:
                                    self.driver.find_element(By.XPATH, "//div[@data-testid='conversation-panel-body'] | //div[contains(@class, 'copyable-area')]")
                                    clic_exitoso = True
                                except:
                                    pass
                        except Exception as e:
                            print(f"   ℹ️ Clic simple falló: {e}")

                    # Método 3: JavaScript click
                    if not clic_exitoso:
                        try:
                            self.driver.execute_script("arguments[0].click();", resultado)
                            print(f"   ✓ Clic con JavaScript")
                            time.sleep(3)

                            # Verificar si se abrió
                            try:
                                self.driver.find_element(By.XPATH, "//header[@data-testid='conversation-header']")
                                clic_exitoso = True
                            except:
                                try:
                                    self.driver.find_element(By.XPATH, "//div[@data-testid='conversation-panel-body'] | //div[contains(@class, 'copyable-area')]")
                                    clic_exitoso = True
                                except:
                                    pass
                        except Exception as e:
                            print(f"   ℹ️ Clic JavaScript falló: {e}")

                    # Verificar resultado final
                    if clic_exitoso:
                        print(f"✅ Comunidad '{nombre_comunidad}' abierta")

                        # IMPORTANTE: Hacer clic en "Detalles del perfil" para abrir el panel de info
                        try:
                            time.sleep(3)
                            print(f"   🔍 Abriendo detalles del perfil...")

                            # Buscar el botón "Detalles del perfil" con el selector exacto
                            boton_detalles = self.wait.until(EC.element_to_be_clickable(
                                (By.XPATH, "//div[@title='Detalles del perfil'][@role='button']")
                            ))
                            boton_detalles.click()
                            print(f"   ✓ Clic en 'Detalles del perfil' exitoso")
                            self.esperar_aleatorio(2, 3)
                        except Exception as e:
                            print(f"   ⚠️ Error abriendo detalles del perfil: {e}")
                            return False

                        self.esperar_aleatorio(2, 3)
                        return True
                    else:
                        print(f"   ⚠️ El chat no se abrió después de 4 intentos")
                        return False
                else:
                    print(f"❌ No se encontró la comunidad '{nombre_comunidad}'")
                    print(f"   💡 Verifica que existe con ese nombre en WhatsApp")
                    return False

            except Exception as e:
                print(f"❌ Error buscando comunidad: {e}")
                return False

        except Exception as e:
            print(f"❌ Error buscando comunidad: {e}")
            return False

    def abrir_info_comunidad(self):
        """Abrir información de la comunidad"""
        try:
            # Hacer clic en el encabezado de la comunidad
            header = self.wait.until(EC.presence_of_element_located(
                (By.XPATH, "//header[@data-testid='conversation-header']")
            ))
            header.click()
            self.esperar_aleatorio(2, 3)
            return True
        except Exception as e:
            print(f"❌ Error abriendo info de comunidad: {e}")
            return False

    def agregar_participante(self, celular):
        """Agregar un participante a la comunidad - PASOS EXACTOS"""
        try:
            # Convertir celular a string y limpiar el .0 si viene de Excel
            celular = str(int(float(celular)))
            print(f"\n➕ Agregando: {celular}")

            # PASO 1: Clic en la comunidad (tab de la comunidad en el panel de info)
            # Selector: div[@role='button'][@data-tab='6'] que contiene el nombre de la comunidad
            try:
                print("  PASO 1: Buscando tab de la comunidad...")
                time.sleep(2)

                # Buscar el botón con data-tab="6" (tab de Avisos/comunidad)
                tab_comunidad = self.wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//div[@role='button'][@data-tab='6']")
                ))
                tab_comunidad.click()
                print("  ✓ Clic en tab de comunidad exitoso")
                self.esperar_aleatorio(2, 3)
            except Exception as e:
                print(f"  ⚠️ Error en PASO 1 (clic en tab comunidad): {e}")
                return False

            # PASO 2: Clic en "Añadir miembros"
            # Selector: button[@aria-label='Añadir miembros'] con icono person-add-filled-refreshed
            try:
                print("  PASO 2: Buscando botón 'Añadir miembros'...")
                time.sleep(1)

                # Buscar por el botón con aria-label exacto
                boton_anadir = self.wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//button[@aria-label='Añadir miembros']")
                ))
                boton_anadir.click()
                print("  ✓ Clic en 'Añadir miembros' exitoso")
                self.esperar_aleatorio(2, 3)
            except Exception as e:
                print(f"  ⚠️ Error en PASO 2 (botón añadir miembros): {e}")
                return False

            # PASO 3: Buscar el contacto en el campo de búsqueda
            # Selector: div[@contenteditable='true'][@data-tab='3'] con aria-label="Buscar un nombre o número"
            try:
                print("  PASO 3: Escribiendo número en campo de búsqueda...")
                time.sleep(1)

                # Buscar campo de búsqueda específico
                campo_busqueda = self.wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//div[@contenteditable='true'][@data-tab='3'][@aria-label='Buscar un nombre o número']")
                ))
                campo_busqueda.click()
                time.sleep(0.5)

                # Limpiar campo
                campo_busqueda.send_keys(Keys.CONTROL + "a")
                campo_busqueda.send_keys(Keys.DELETE)
                time.sleep(0.5)

                # Escribir el número con prefijo +57
                celular_completo = f"+57{celular}"
                campo_busqueda.send_keys(celular_completo)
                print(f"  ✓ Escrito: {celular_completo}")
                time.sleep(2)
            except Exception as e:
                print(f"  ⚠️ Error en PASO 3 (escribir número): {e}")
                return False

            # PASO 4: Presionar Enter para buscar
            try:
                print("  PASO 4: Presionando Enter...")
                campo_busqueda.send_keys(Keys.ENTER)
                print("  ✓ Enter presionado")
                self.esperar_aleatorio(3, 4)
            except Exception as e:
                print(f"  ⚠️ Error en PASO 4 (Enter): {e}")
                return False

            # PASO 5: Clic en el botón de confirmar (checkmark)
            # Selector: div[@role='button'] con span[@data-icon='checkmark-medium']
            try:
                print("  PASO 5: Buscando botón de confirmar (checkmark)...")
                time.sleep(2)

                # Buscar el botón con el icono de checkmark
                boton_checkmark = self.wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//span[@data-icon='checkmark-medium']/ancestor::div[@role='button'][1]")
                ))
                boton_checkmark.click()
                print("  ✓ Clic en checkmark exitoso")
                self.esperar_aleatorio(2, 3)
            except Exception as e:
                print(f"  ⚠️ Error en PASO 5 (checkmark): {e}")
                return False

            # PASO 6: Clic en "Añadir miembro" final
            # Selector: div[@role='button'] que contiene span con texto "Añadir miembro"
            try:
                print("  PASO 6: Buscando botón final 'Añadir miembro'...")
                time.sleep(2)

                # Buscar el botón con las clases específicas y el texto "Añadir miembro"
                boton_confirmar = self.wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//div[contains(@class, 'x1i10hfl') and contains(@class, 'x1qjc9v5')]//span[contains(text(), 'Añadir miembro')]")
                ))

                # Intentar clic normal, si falla usar JavaScript
                try:
                    boton_confirmar.click()
                except:
                    self.driver.execute_script("arguments[0].click();", boton_confirmar)

                print("  ✓ Clic en 'Añadir miembro' final exitoso")
                print(f"✅ Participante {celular} agregado exitosamente")
                self.esperar_aleatorio(2, 3)

                # Cerrar ventanas
                try:
                    ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
                    time.sleep(1)
                except:
                    pass

                return True

            except Exception as e:
                print(f"  ⚠️ Error en PASO 6 (botón final añadir): {e}")

                # Intentar cerrar
                try:
                    ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
                    time.sleep(1)
                except:
                    pass
                return False

        except Exception as e:
            print(f"❌ Error general agregando {celular}: {e}")
            return False

    def eliminar_participante(self, celular):
        """Eliminar un participante de la comunidad - PASOS EXACTOS ACTUALIZADOS"""
        try:
            # Convertir celular a string y limpiar el .0 si viene de Excel
            celular = str(int(float(celular)))
            print(f"\n➖ Eliminando: {celular}")

            time.sleep(2)

            # PASO 1: Clic en el tab "Comunidad"
            # Selector: button[@role='tab'] con title="Comunidad"
            try:
                print("  PASO 1: Haciendo clic en tab 'Comunidad'...")
                time.sleep(2)

                # Buscar el botón tab "Comunidad"
                tab_comunidad = self.wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//button[@role='tab' and @title='Comunidad']")
                ))
                tab_comunidad.click()
                print("  ✓ Clic en tab 'Comunidad' exitoso")
                print("  ⏳ Esperando que cargue la vista de comunidad...")
                # Esperar más tiempo porque la vista de comunidad se demora en cargar
                self.esperar_aleatorio(4, 6)
            except Exception as e:
                print(f"  ⚠️ Error en PASO 1 (tab comunidad): {e}")
                return False

            # PASO 2: Clic en "X miembros de la comunidad" (el botón con ícono de búsqueda)
            # Este es el div con role="button" que contiene el texto de miembros y el ícono search
            try:
                print("  PASO 2: Haciendo clic en 'miembros de la comunidad'...")
                time.sleep(2)

                # Buscar el botón que contiene "miembros de la comunidad" y el ícono search
                boton_miembros = None

                # Método 1: Por el ícono search dentro de un botón que tiene el texto "miembros"
                try:
                    boton_miembros = self.wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//div[@role='button' and contains(@class, 'x1ypdohk')]//span[@data-icon='search']/..")
                    ))
                    print("  ✓ Botón 'miembros' encontrado (método 1)")
                except:
                    pass

                # Método 2: Por el div que contiene el span con "miembros de la comunidad"
                if not boton_miembros:
                    try:
                        # Buscar el span que contiene "miembros de la comunidad" y obtener el div padre clickeable
                        span_miembros = self.driver.find_element(
                            By.XPATH,
                            "//span[contains(text(), 'miembros de la comunidad')]"
                        )
                        boton_miembros = span_miembros.find_element(By.XPATH, "./ancestor::div[@role='button'][1]")
                        print("  ✓ Botón 'miembros' encontrado (método 2)")
                    except:
                        pass

                if boton_miembros:
                    boton_miembros.click()
                    print("  ✓ Clic en 'miembros de la comunidad' exitoso")
                    self.esperar_aleatorio(2, 3)
                else:
                    print("  ⚠️ No se encontró el botón de miembros")
                    return False

            except Exception as e:
                print(f"  ⚠️ Error en PASO 2 (botón miembros): {e}")
                return False

            # PASO 3: Escribir el celular en el campo "Buscar miembros"
            # Selector: div[@aria-label="Buscar miembros"][@contenteditable="true"]
            try:
                print("  PASO 3: Escribiendo número en campo 'Buscar miembros'...")
                time.sleep(2)

                # Buscar el campo por aria-label="Buscar miembros"
                campo_busqueda = None

                # Método 1: Por aria-label exacto
                try:
                    campo_busqueda = self.wait.until(EC.presence_of_element_located(
                        (By.XPATH, "//div[@aria-label='Buscar miembros' and @contenteditable='true']")
                    ))
                    print("  ✓ Campo 'Buscar miembros' encontrado (método 1)")
                except:
                    pass

                # Método 2: Buscar el <p> hijo dentro del div con aria-label
                if not campo_busqueda:
                    try:
                        campo_busqueda = self.driver.find_element(
                            By.XPATH,
                            "//div[@aria-label='Buscar miembros']//p[contains(@class, 'selectable-text')]"
                        )
                        print("  ✓ Campo encontrado (método 2: p dentro del div)")
                    except:
                        pass

                if campo_busqueda:
                    campo_busqueda.click()
                    time.sleep(0.5)

                    # Limpiar campo
                    campo_busqueda.send_keys(Keys.CONTROL + "a")
                    campo_busqueda.send_keys(Keys.DELETE)
                    time.sleep(0.5)

                    # Escribir el número con prefijo +57
                    celular_completo = f"+57{celular}"
                    campo_busqueda.send_keys(celular_completo)
                    print(f"  ✓ Escrito: {celular_completo}")
                    self.esperar_aleatorio(2, 3)
                else:
                    print("  ⚠️ No se encontró el campo 'Buscar miembros'")
                    return False

            except Exception as e:
                print(f"  ⚠️ Error en PASO 3 (escribir en buscar miembros): {e}")
                return False

            # PASO 4: Hacer clic en el resultado (el contacto encontrado)
            try:
                print("  PASO 4: Haciendo clic en el contacto encontrado...")
                time.sleep(2)

                # Buscar el primer resultado con la clase específica
                contacto = self.wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//div[contains(@class, '_ak8l') and contains(@class, '_ap1_')]")
                ))
                contacto.click()
                print("  ✓ Clic en contacto exitoso")
                self.esperar_aleatorio(2, 3)
            except Exception as e:
                print(f"  ⚠️ Error en PASO 4 (clic en contacto): {e}")
                return False

            # PASO 5: Clic en "Eliminar de la comunidad"
            # Selector: div que contiene el SVG close-circle-refreshed y el span con texto "Eliminar de la comunidad"
            try:
                print("  PASO 5: Buscando opción 'Eliminar de la comunidad'...")
                time.sleep(2)

                # Método 1: Por el span con el texto y clases específicas
                opcion_eliminar = None
                try:
                    opcion_eliminar = self.wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//span[contains(@class, 'x1o2sk6j') and contains(text(), 'Eliminar de la comunidad')]")
                    ))
                    print("  ✓ Opción eliminar encontrada (método 1: span texto)")
                except:
                    pass

                # Método 2: Por el div padre que contiene el icono close-circle-refreshed
                if not opcion_eliminar:
                    try:
                        # Buscar el div que contiene el SVG con title="close-circle-refreshed"
                        div_eliminar = self.driver.find_element(
                            By.XPATH,
                            "//svg[@data-icon='close-circle-refreshed']/ancestor::div[contains(@class, 'x1c4vz4f')][1]"
                        )
                        opcion_eliminar = div_eliminar
                        print("  ✓ Opción eliminar encontrada (método 2: div con icono)")
                    except:
                        pass

                if opcion_eliminar:
                    opcion_eliminar.click()
                    print("  ✓ Clic en 'Eliminar de la comunidad' exitoso")
                    self.esperar_aleatorio(2, 3)
                else:
                    print("  ⚠️ No se encontró la opción 'Eliminar de la comunidad'")
                    return False

            except Exception as e:
                print(f"  ⚠️ Error en PASO 5 (opción eliminar): {e}")
                return False

            # PASO 6: Confirmar eliminación haciendo clic en el botón "Eliminar"
            # Selector: span con texto "Eliminar" y clases específicas
            try:
                print("  PASO 6: Confirmando eliminación con botón 'Eliminar'...")
                time.sleep(2)

                # Buscar el span con el texto "Eliminar" exacto
                boton_confirmar = self.wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//span[contains(@class, 'x140p0ai') and text()='Eliminar']")
                ))

                # Intentar clic normal, si falla usar JavaScript
                try:
                    boton_confirmar.click()
                except:
                    self.driver.execute_script("arguments[0].click();", boton_confirmar)

                print("  ✓ Clic en botón 'Eliminar' confirmado")
                print(f"✅ Participante {celular} eliminado exitosamente")
                self.esperar_aleatorio(2, 3)

                # Cerrar ventanas modales
                self._cerrar_ventanas_modales()

                return True

            except Exception as e:
                print(f"  ⚠️ Error en PASO 6 (confirmar eliminar): {e}")
                self._cerrar_ventanas_modales()
                return False

        except Exception as e:
            print(f"❌ Error general eliminando {celular}: {e}")
            return False

    def procesar_excel(self):
        """Procesar archivo Excel con las listas"""
        try:
            # Buscar archivo Excel
            archivos_excel = [f for f in os.listdir('.') if f.endswith('.xlsx') and 'comunidades' in f.lower()]

            if not archivos_excel:
                print("❌ No se encontró archivo Excel")
                print("💡 Debe existir un archivo que contenga 'comunidades' en el nombre")
                return False

            archivo = archivos_excel[0]
            print(f"\n📖 Procesando: {archivo}")

            # Cargar Excel
            df = pd.read_excel(archivo)

            # Verificar columnas
            columnas_requeridas = ['Comunidad_Agregar', 'Celular_Agregar', 'Comunidad_Eliminar', 'Celular_Eliminar']

            print(f"\n📋 Columnas encontradas: {list(df.columns)}")

            # Limpiar datos
            df = df.fillna('')

            # Usar la cantidad configurada al inicio
            print(f"\n🔢 Total de registros en Excel: {len(df)}")

            if self.cantidad_procesar is None:
                # Procesar todos
                df_procesar = df
                print(f"🚀 Procesando TODOS los {len(df_procesar)} registros...")
            else:
                # Procesar cantidad específica
                df_procesar = df.head(self.cantidad_procesar)
                print(f"🚀 Procesando {len(df_procesar)} registros (de {len(df)} totales)...")

            # Estadísticas
            agregados_ok = 0
            agregados_error = 0
            eliminados_ok = 0
            eliminados_error = 0

            # Procesar cada fila
            for i, row in df_procesar.iterrows():
                print(f"\n{'='*60}")
                print(f"📊 Procesando registro {i+1}/{len(df_procesar)}")
                print(f"{'='*60}")

                # PROCESO 1: AGREGAR
                if row['Comunidad_Agregar'] and row['Celular_Agregar']:
                    comunidad_agregar = str(row['Comunidad_Agregar']).strip()
                    celular_agregar = str(row['Celular_Agregar']).strip()

                    print(f"\n➕ PROCESO: AGREGAR")
                    print(f"   Comunidad: {comunidad_agregar}")
                    print(f"   Celular: {celular_agregar}")

                    if self.buscar_comunidad(comunidad_agregar):
                        if self.agregar_participante(celular_agregar):
                            agregados_ok += 1
                        else:
                            agregados_error += 1

                        # Cerrar cualquier ventana abierta y volver a la vista principal
                        self._cerrar_ventanas_modales()

                    # Esperar entre procesos
                    print(f"\n⏳ Esperando {self.tiempo_entre_procesos} segundos antes del siguiente proceso...")
                    time.sleep(self.tiempo_entre_procesos)

                # PROCESO 2: ELIMINAR
                if row['Comunidad_Eliminar'] and row['Celular_Eliminar']:
                    comunidad_eliminar = str(row['Comunidad_Eliminar']).strip()
                    celular_eliminar = str(row['Celular_Eliminar']).strip()

                    print(f"\n➖ PROCESO: ELIMINAR")
                    print(f"   Comunidad: {comunidad_eliminar}")
                    print(f"   Celular: {celular_eliminar}")

                    if self.buscar_comunidad(comunidad_eliminar):
                        if self.eliminar_participante(celular_eliminar):
                            eliminados_ok += 1
                        else:
                            eliminados_error += 1

                        # Cerrar cualquier ventana abierta y volver a la vista principal
                        self._cerrar_ventanas_modales()

                # Esperar entre contactos (si no es el último)
                if i < len(df_procesar) - 1:
                    self.esperar_aleatorio(self.tiempo_min_contacto, self.tiempo_max_contacto)

            # Mostrar estadísticas finales
            print("\n" + "="*60)
            print("📊 ESTADÍSTICAS FINALES")
            print("="*60)
            print(f"➕ Agregados exitosos: {agregados_ok}")
            print(f"❌ Errores al agregar: {agregados_error}")
            print(f"➖ Eliminados exitosos: {eliminados_ok}")
            print(f"❌ Errores al eliminar: {eliminados_error}")
            print(f"📈 Total procesados: {agregados_ok + agregados_error + eliminados_ok + eliminados_error}")
            print("="*60)

            return True

        except Exception as e:
            print(f"❌ Error procesando Excel: {e}")
            return False

    def ejecutar(self):
        """Ejecutar proceso completo"""
        try:
            print("\n" + "="*60)
            print("🤖 GESTOR DE COMUNIDADES DE WHATSAPP")
            print("="*60)

            # Configurar parámetros
            self.configurar_parametros()

            # Configurar navegador
            if not self.configurar_navegador():
                return

            # Iniciar WhatsApp
            if not self.iniciar_whatsapp():
                return

            # Procesar Excel
            self.procesar_excel()

            print("\n🎉 ¡Proceso completado!")

        except Exception as e:
            print(f"❌ Error general: {e}")
        finally:
            if self.driver:
                input("\n⏸️ Presiona Enter para cerrar el navegador...")
                self.driver.quit()


if __name__ == "__main__":
    gestor = GestorComunidadesWhatsApp()
    gestor.ejecutar()
