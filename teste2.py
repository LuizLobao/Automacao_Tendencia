from playwright.sync_api import Playwright, sync_playwright
import time

def send_message_to_group(group_name: str, message: str):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        
        # Abra o site do WhatsApp Web e faça login manualmente
        page.goto('https://web.whatsapp.com')
        time.sleep(15)
        
        # Selecione o grupo para o qual deseja enviar a mensagem
        group_selector = f"span[title='{group_name}']"
        group = page.wait_for_selector(group_selector)
        group.click()
        
        # Digite a mensagem
        message_box = page.wait_for_selector('div[data-tab="6"]')
        message_box.click()
        message_box.type(message)
        
        # Envie a mensagem
        message_box.press('Enter')
        
        # Feche o navegador
        browser.close()

send_message_to_group("Teste Python", "Sua mensagem aqui")
