# Automação para consulta de CPF

Uma automação usando Python, Selenium e PyOpenXL focara em consultar o CPF do cliente e conferir qual status de pagamento do mesmo.
Caso o pagamento estiver "atrasado", será marcado como pendente em uma nova planilha.
Caso o pagamento estiver "em dia", será marcado como status "em dia" e iremos extrair a data e o metodo de pagamento do cliente para nova planilha
