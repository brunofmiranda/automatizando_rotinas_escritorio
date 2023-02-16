from shareplum import Site    
from shareplum import Office365    
from shareplum.site import Version
import sharepy

#conecta no sharepoint via shareplum
def connectSharepoint(user, psswrd, organization_site, my_site):     
    authcookie = Office365(organization_site, username = user, password = psswrd).GetCookies()    
    return Site(my_site, version=Version.v2016, authcookie=authcookie)

if __name__ == '__main__':

    #credenciais e sharepoint a ser utilizado
    user = 'usuario@mulherPepita.com'
    password = 'senhaIndecifrável'
    organization_site = 'https://empresa.sharepoint.com'
    my_site = 'https://empresa.sharepoint.com/sites/nome_site'

    #conexões - shareplum e sharepy
    site = connectSharepoint(user, password, organization_site, my_site)
    s = sharepy.connect(organization_site, user, password)

    #pegando a lista e as colunas desejadas com o shareplum
    sp_list = site.List('NOME DA LISTA')
    sp_data = sp_list.GetListItems(fields=['COLUNAS', 'DA LISTA'])

    #iterando sobre cada linha
    for row in sp_data:

        #nome do arquivo
        file_name = row['Coluna para ser utilizada como nome do arquivo']

        #tentar pegar o arquivo
        try:
             
            #coluna do arquivo é um json em string
            dict_file = eval(row['Coluna do arquivo'])
            #puxando o caminho do arquivo
            file_patch = dict_file['serverRelativeUrl']
            #fazendo download do arquivo
            r = s.getfile(organization_site + file_patch, filename = 'arquivos/' + file_name + '.jpeg')
        except:
            
            #escrevendo a inscrição do arquivo que deu erro
            meuArquivo = open('file.txt', 'a')
            meuArquivo.write(f'{file_name}\n')
            meuArquivo.close()