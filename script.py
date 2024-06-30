import time
from pypresence import Presence
import win32gui
import win32process
import psutil

# ID de la aplicación de Discord
client_id = "1257074430891393156"

# Función para obtener la ventana activa de Word
def get_word_window_info():
    word_class = 'OpusApp'
    def enum_windows_callback(hwnd, windows):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetClassName(hwnd) == word_class:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if psutil.Process(pid).name().upper() == 'WINWORD.EXE':
                windows.append(hwnd)
        return True
    windows = []
    win32gui.EnumWindows(enum_windows_callback, windows)
    return windows

# Función para obtener el nombre del documento de Word
def get_word_doc_name(hwnd):
    return win32gui.GetWindowText(hwnd)

# Función principal para actualizar el estado en Discord
def main():
    try:
        RPC = Presence(client_id)
        RPC.connect()
    except Exception as e:
        print(f"Error connecting to Discord: {e}")
        return

    start_time = time.time()

    # Imágenes y detalles predeterminados cuando no hay documento abierto
    default_details = "Idle"
    default_state = "No document open"
    default_large_image = "word_icon"  # Nombre de la imagen predeterminada cargada en Rich Presence
    default_large_text = "Idle"
    
    while True:
        try:
            word_window = get_word_window_info()
            if word_window:
                hwnd = word_window[0]
                doc_name = get_word_doc_name(hwnd)
                RPC.update(
                    details="Working on a document",
                    state=f"Editing: {doc_name}",
                    large_image="wrp_open",  # Nombre de la imagen cargada en Rich Presence
                    large_text="Microsoft Word",
                    small_image="word_icon",  # Nombre de la imagen cargada en Rich Presence
                    small_text="Editing",
                    start=start_time
                )
            else:
                RPC.update(
                    details=default_details,
                    state=default_state,
                    large_image=default_large_image,
                    large_text=default_large_text,
                    small_image=None,  # Puedes omitir esto si no deseas mostrar una imagen pequeña
                    small_text=None
                )
        except Exception as e:
            print(f"Error updating Discord presence: {e}")

        time.sleep(15)  # Espera 15 segundos antes de verificar nuevamente

if __name__ == "__main__":
    main()
