from functions import generate_orders_batch
from functions import send_email

if __name__ == "__main__":
    generate_orders_batch("pedidos.csv")

    send_email(
    to="bruno_rdss@hotmail.com",
    cc="rafake_9@hotmail.com;bruno.-.santos@hotmail.com",
    bcc="bruno.-.santos@hotmail.com",
    subject="Order Request",
    html_body="""<p>Dear all,</p>
                 <p>Please note it attached.</p>
                 <img src='seu_link_imagem' wdth=200>"""
)