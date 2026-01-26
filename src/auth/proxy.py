# import httplib2
# import requests


# # PREENCHA AQUI SEUS DADOS DE PROXY
# PROXY_HOST = "seu.proxy.com.br"
# PROXY_PORT = 8080
# PROXY_USER = "seu_usuario"
# PROXY_PASS = "sua_senha"

# def get_google_http_client():
#     """
#     Retorna o cliente HTTP (httplib2) configurado com Proxy.
#     Use na função 'build()' do Google Sheets.
#     """
#     proxy_info = httplib2.ProxyInfo(
#         httplib2.socks.PROXY_TYPE_HTTP,
#         PROXY_HOST,
#         PROXY_PORT,
#         proxy_user=PROXY_USER,
#         proxy_pass=PROXY_PASS
#     )
#     return httplib2.Http(proxy_info=proxy_info)

# def get_requests_session():
#     """
#     Retorna uma Sessão Requests configurada com Proxy.
#     Use na função 'creds.refresh()' na autenticação.
#     """
#     session = requests.Session()
#     # Monta a URL do proxy: http://user:pass@host:port
#     proxy_url = f"http://{PROXY_USER}:{PROXY_PASS}@{PROXY_HOST}:{PROXY_PORT}"
    
#     session.proxies = {
#         "http": proxy_url,
#         "https": proxy_url
#     }
#     return session