from playwright.sync_api import sync_playwright
import time

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()
    page.goto('http://localhost:8505', timeout=15000)
    time.sleep(5)

    print("Página carregada")

    # Preenche Técnico (primeiro text input)
    text_inputs = page.query_selector_all('input[type="text"]')
    print(f"Text inputs encontrados: {len(text_inputs)}")
    if text_inputs:
        text_inputs[0].fill("Joao Teste")
        text_inputs[0].press("Tab")
    time.sleep(0.5)

    # Preenche Município (segundo text input ou selectbox)
    if len(text_inputs) > 1:
        text_inputs[1].fill("Campo Grande")
        text_inputs[1].press("Tab")
    time.sleep(0.5)

    # Preenche Área (primeiro number input)
    num_inputs = page.query_selector_all('input[type="number"]')
    print(f"Number inputs encontrados: {len(num_inputs)}")
    if num_inputs:
        num_inputs[0].triple_click()
        num_inputs[0].fill("100")
    time.sleep(0.3)

    # Preenche Produtividade
    if len(num_inputs) > 1:
        num_inputs[1].triple_click()
        num_inputs[1].fill("60")
    time.sleep(0.3)

    # Clica no botão CALCULAR E SALVAR
    btn = page.query_selector('button:has-text("CALCULAR")')
    print(f"Botão CALCULAR encontrado: {btn is not None}")
    if btn:
        btn.click()
        time.sleep(4)

    # Captura conteúdo após clique
    content = page.inner_text("body")
    print("\n=== RESULTADO APÓS CLIQUE ===")
    for line in content.split("\n"):
        line = line.strip()
        if line and len(line) > 2:
            print(repr(line[:100]))

    # Verifica se há erros
    errors = page.query_selector_all('[data-testid="stAlert"]')
    print(f"\nAlertas na página: {len(errors)}")
    for e in errors:
        print("  ALERTA:", e.inner_text()[:80])

    browser.close()
    print("\nTeste concluído")
