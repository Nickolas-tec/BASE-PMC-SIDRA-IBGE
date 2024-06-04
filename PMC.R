
# CARREGAMENTO DAS BIBLITOECAS

library(dplyr)
library(sidrar)
library(stringr)
library(tidyverse)
library(tidyr)
library(zoo)
library(openxlsx)



############################################################# PMC/SIDRA/TRATAMENTO/UNIFICANDO ####################################################################################################

######################## EXTRAINDO E TRATANDO NUMERO INDICE 8880/8881/8883 ########################################################################################################################  


# PLANILHA 1- NUMERO INDICE

volume_vendas_varejista = get_sidra(api = "/t/8880/n1/all/v/7169/p/all/c11046/56734/d/v7169%205")%>% 
  select(`Mês (Código)`, `Mês`, `Valor`)
names(volume_vendas_varejista)[3] <- "volume_de_vendas_varejista"


volume_vendas_ampliado = get_sidra(api = "/t/8881/n1/all/v/7169/p/all/c11046/56736/d/v7169%205")%>%
  select(Valor,`Mês (Código)`)
names(volume_vendas_ampliado)[1] <- "volume_vendas_ampliado"


volume_vendas_ampliado_atividades = get_sidra(api = "/t/8883/n1/all/v/7169/p/all/c11046/56736/c85/all/d/v7169%205")%>%
  select(Valor, Atividades, `Mês (Código)`)


# PIVOTANDO COLUNA ATIVIDADES DO DF VOLUME_VENDAS_ATIVIDADES
# E CRIANDO O DATAFRAME VOLUME_VENDAS

# DATAFRAME PIVOTADO

df1 <- pivot_wider(volume_vendas_ampliado_atividades , names_from = Atividades, values_from = c("Valor"))


# CRIANDO DATAFRAME JJ PARA PODER FAZER O MERGE SEM CONFLITO DE NUMEROS DE LINHAS

jj <- volume_vendas_varejista[ ,c("Mês (Código)" ,
                                  "volume_de_vendas_varejista")]


# MERGE DOS DATAFRAMES

volume_vendas <- merge(jj,volume_vendas_ampliado, by ="Mês (Código)") 

# MERGE DO MERGE

numero_indice <- merge(volume_vendas,df1, by = "Mês (Código)") 


# CRIANDO COLUNA DATA EM VARIACAO_MES_ANTERIOR

numero_indice <- numero_indice %>% dplyr::mutate(Data = as.Date(paste0(substr(numero_indice$`Mês (Código)`, 
                                                                              start = 1, stop = 4),"-",
                                                                       substr(numero_indice$`Mês (Código)`, 
                                                                              start = 5 , stop = 6), "-01")))



# ALOCANDO DATA NA PRIMEIRA COLUNA

numero_indice <- numero_indice %>%
  select(Data, everything())



#### RENOMEANDO AS COLUNA DO DATAFRAME

colnames(numero_indice) <- c("Data",
                             "mes",
                             "volume_de_vendas_varejista",
                             "volume_de_vendas_ampliado",
                             "combustíveis_e_lubrificantes",
                             "hipermer_supermer_prod_alimen_bebidas",
                             "hipermercados_supermercados",
                             "tecidos_vestuário_calçados",
                             "moveis_eletrodomesticos",
                             "moveis",
                             "eletrodomesticos",
                             "artigos_farmacêuticos_medicos_cosmeticos",
                             "livros_jornais_papelaria",
                             "equipamentos_materiais_escritorio_informatica",
                             "outros_artigos",
                             "veiculos_motocicletas",
                             "material_de_construcao",
                             "atacado_especializado")





######################################################################################################################################################################################################################


# PLANILHA 2 - INDICE AJUSTE SAZONAL CALCULANDO A MEDIA 

#          EXTRAINDO DADOS E TRATANDO AJUSTE SAZONAL 8880/8881/8883   


indice_ajuste_sazonal_varejista = get_sidra(api = "/t/8880/n1/all/v/7170/p/all/c11046/56734/d/v7170%205")%>%
  select(Mês, `Mês (Código)`, Valor)
#mutate(Valor = as.numeric(Valor))
names(indice_ajuste_sazonal_varejista)[3] <- "indice_ajuste_sazonal_varejista"


indice_ajuste_sazonal_ampliado = get_sidra(api = "/t/8881/n1/all/v/7170/p/all/c11046/56736/d/v7170%205")%>%
  select(Valor, `Mês (Código)`)
#mutate(Valor = as.numeric(Valor))
names(indice_ajuste_sazonal_ampliado)[1] <- "indice_ajuste_sazonal_ampliado"



indice_ajuste_ampliado_atividades = get_sidra(api = "/t/8883/n1/all/v/7170/p/all/c11046/56736/c85/all/d/v7170%205")%>%
  select(Valor, Atividades, `Mês (Código)`)
#mutate(Valor = as.numeric(Valor))



# PIVOTANDO INIDCE_AJUSTE_AMPLIADO_ATIVIDADES E GERANDO UM NOVO DATAFRAME

df2 <- pivot_wider(indice_ajuste_ampliado_atividades, names_from = Atividades, values_from = c("Valor"))


# CRIANDO DATAFRAME BB


bb <- indice_ajuste_sazonal_varejista[,c("Mês (Código)","indice_ajuste_sazonal_varejista")]

# MERGE DOS DATAFRAMES

indice_ajuste_sazonal <- merge(bb,indice_ajuste_sazonal_ampliado, by = "Mês (Código)")


# MERGE DO MERGE

indice_ajuste_sazonal <- merge(indice_ajuste_sazonal, df2, by = "Mês (Código)")





# CRIANDO COLUNA DATA EM INDICE_AJUSTE_SAZONAL


indice_ajuste_sazonal <- indice_ajuste_sazonal %>% dplyr::mutate(Data = as.Date(paste0(substr(indice_ajuste_sazonal$`Mês (Código)`, start = 1, stop = 4),"-",
                                                                                       substr(indice_ajuste_sazonal$`Mês (Código)`, 
                                                                                              start = 5 , stop = 6), "-01")))
# ALOCANDO DATA NA PRIMEIRA COLUNA

indice_ajuste_sazonal <- indice_ajuste_sazonal %>%
  select(Data, everything())


# RENOMEANDO AS COLUNAS

colnames(indice_ajuste_sazonal) <- c("Data",
                                     "mes",
                                     "indice_ajuste_sazonal_varejista",
                                     "indice_ajuste_sazonal_ampliado",
                                     "combustíveis_e_lubrificantes",
                                     "hipermer_supermer_prod_alimen_bebidas",
                                     "hipermercados_supermercados",
                                     "tecidos_vestuário_calçados",
                                     "moveis_eletrodomesticos",
                                     "moveis",
                                     "eletrodomesticos",
                                     "artigos_farmacêuticos_medicos_cosmeticos",
                                     "livros_jornais_papelaria",
                                     "equipamentos_materiais_escritorio_informatica",
                                     "outros_artigos",
                                     "veiculos_motocicletas",
                                     "material_de_construcao",
                                     "atacado_especializado")



numero_indice_ajuste_sazonal <- indice_ajuste_sazonal



############## CALCULANDO AS MEDIAS ############

# CALCULO TRIMESTRE MOVEL VAREJISTA

# MEDIA MOVEL/VAREJISTA

indice_ajuste_sazonal <- indice_ajuste_sazonal %>%
  
  mutate(media_movel_varejista = (lag(indice_ajuste_sazonal_varejista,
                                      
                                      2) + lag(indice_ajuste_sazonal_varejista, 1)
                                  
                                  + indice_ajuste_sazonal_varejista) / 3)



# CALCULO TRIMESTRE MOVEL AMPLIADO

# MEDIA MOVEL/AMPLIADO

indice_ajuste_sazonal <- indice_ajuste_sazonal %>%
  mutate(media_movel_ampliado = (lag(indice_ajuste_sazonal_ampliado,
                                     2) + lag(indice_ajuste_sazonal_ampliado, 1)
                                 + indice_ajuste_sazonal_ampliado) / 3)




# CALCULO TRIMESTRE MOVEL COMBUSTIVEIS E LUBRIFICANTES


# MEDIA MOVEL/COMBUSTIVEIS E LUBRIFICANTES = COMBLUB


indice_ajuste_sazonal <- indice_ajuste_sazonal %>%
  mutate(medial_movel_COMBLUB = (lag(`combustíveis_e_lubrificantes`,
                                     2) + lag(`combustíveis_e_lubrificantes`, 1)
                                 + `combustíveis_e_lubrificantes`) / 3)






# CALCULO TRIMESTRE MOVEL SUPERMERCADOS. HIPERMERCADOS, BEBIDAS E FUMO


# MEDIA MOVEL/ HIPERMERCADOS, SUPERMERCADOS, PRODUTOS ALIMENTICIOS, BEBIDAS E FUMO = HSPBF

indice_ajuste_sazonal <- indice_ajuste_sazonal %>%
  mutate(media_movel_HSPBF = (lag(`hipermer_supermer_prod_alimen_bebidas`,
                                  2) + lag(`hipermer_supermer_prod_alimen_bebidas`, 1)
                              + `hipermer_supermer_prod_alimen_bebidas`) / 3)




# CALCULO TRIMESTRE MOVEL HIPERMERCADOS E SUPERMERCADOS


# MEDIA MOVEL/ HIPERMERCADOS E SUPERMERCADOS = HS

indice_ajuste_sazonal <- indice_ajuste_sazonal %>%
  mutate(media_movel_HS = (lag(`hipermercados_supermercados`,
                               2) + lag(`hipermercados_supermercados`, 1)
                           + `hipermercados_supermercados`) / 3)



# CALCULO TRIMESTRE MOVEL TECIDOS,VESTUARIO E CALÇADOS


# MEDIA MOVEL/ TECIDOS, VESTUARIO E CALÇADOS = TVC

indice_ajuste_sazonal <- indice_ajuste_sazonal %>%
  mutate(media_movel_TVC = (lag(`tecidos_vestuário_calçados`,
                                2) + lag(`tecidos_vestuário_calçados`, 1)
                            + `tecidos_vestuário_calçados`) / 3)



#  CALCULO TRIMESTRE MOVEL MOVEIS E ELETRODOMESTICOS

# MEDIA MOVEL/ MOVEIS E ELETRODOMESTICOS = ME

indice_ajuste_sazonal <- indice_ajuste_sazonal %>%
  mutate(media_movel_ME = (lag(`moveis_eletrodomesticos`,
                               2) + lag(`moveis_eletrodomesticos`, 1)
                           + `moveis_eletrodomesticos`) / 3)



# CALCULO TRIMESTRE MOVEL MOVEIS

# MEDIA MOVEL/ MOVEIS

indice_ajuste_sazonal <- indice_ajuste_sazonal %>%
  mutate(media_movel_moveis = (lag(moveis,
                                   2) + lag(moveis, 1)
                               + moveis) / 3)



# CALCULO TRIMESTRE MOVEL ELETRODOMESTICOS

# MEDIA MOVEL/ ELETRODOMESTICOS

indice_ajuste_sazonal <- indice_ajuste_sazonal %>%
  mutate(media_movel_eletrodomesticos = (lag(eletrodomesticos,
                                             2) + lag(eletrodomesticos, 1)
                                         + eletrodomesticos) / 3)



# CALCULO TRIMESTRE MOVEL ARTIFGOS FARMACEUTICOS


# MEDIA MOVEL/ ARTIGOS FARMACEUTICOS, MEDICOS E ORTOPEDICOS DE PERFUMARIA E COSMETICOS = AMOPC

indice_ajuste_sazonal <- indice_ajuste_sazonal %>%
  mutate(media_movel_AMOPC = (lag(`artigos_farmacêuticos_medicos_cosmeticos`,
                                  2) + lag(`artigos_farmacêuticos_medicos_cosmeticos`, 1)
                              + `artigos_farmacêuticos_medicos_cosmeticos`) / 3)




# CALCULO TRIMESTRE MOVEL LIVROS,JORNAIS, REVISTAS E PAPELARA


# MEDIA MOVEL / LIVROS, JORNAIS, REVISTAS E PAPELARIA = LJRP

indice_ajuste_sazonal <- indice_ajuste_sazonal %>%
  mutate(media_movel_LJRP = (lag(`livros_jornais_papelaria`,
                                 2) + lag(`livros_jornais_papelaria`, 1)
                             + `livros_jornais_papelaria`) / 3)





# CALCULO TRIMESTRE MOVEL EQUIPAMENTOS E MATERIAIS PARA ESCRITORIO, INFORMATICA E COMUNICAÇÃO 



# MEDIA MOVEL / EQUIPAMENTOS E MATERIAIS PARA ESCRITORIO, INFORMATICA E COMUNICAÇÃO = EMEIC

indice_ajuste_sazonal <- indice_ajuste_sazonal %>%
  mutate(media_movel_EMEIC = (lag(`equipamentos_materiais_escritorio_informatica`,
                                  2) + lag(`equipamentos_materiais_escritorio_informatica`, 1)
                              + `equipamentos_materiais_escritorio_informatica`) / 3)





# CALCULO TRIMESTRE MOVELUSO PESSOAL DOMESTICO

# MEDIA MOVEL/ OUTROS ARTIGOS DE USO PESSOAL E DOMESTICO = OAPD

indice_ajuste_sazonal <- indice_ajuste_sazonal %>%
  mutate(media_movel_OAPD = (lag(`outros_artigos`,
                                 2) + lag(`outros_artigos`, 1)
                             + `outros_artigos`) / 3)




# CALCULO TRIMESTRE MOVEL VEICULOS MOTOCICLETAS

# MEDIA MOVEL/ VEICULOS, MOTOCICLETAS, PARTES E PEÇAS = VMPP

indice_ajuste_sazonal <- indice_ajuste_sazonal %>%
  mutate(media_movel_VMPP = (lag(`veiculos_motocicletas`,
                                 2) + lag(`veiculos_motocicletas`, 1)
                             + `veiculos_motocicletas`) / 3)





# CALCULO TRIMESTRE MOVEL MATERIAL DE CONSTRUÇÃO

# MEDIA MOVEL/ MATERIAIS DE CONSTRUÇÃO = MC

indice_ajuste_sazonal <- indice_ajuste_sazonal %>%
  mutate(media_movel_MC = (lag(`material_de_construcao`,
                               2) + lag(`material_de_construcao`, 1)
                           + `material_de_construcao`) / 3)







# CALCULO TRIMESTRE MOVEL ATACADO ESPECIALIZADO

# MEDIA MOVEL/ ATACADO ESPECIALIZADO EM PRODUTOS ALIMENTICIOS, BEBIDAS E FUMO = AEPABF

indice_ajuste_sazonal <- indice_ajuste_sazonal %>%
  mutate(media_movel_AEPABF = (lag(`atacado_especializado`,
                                   2) + lag(`atacado_especializado`, 1)
                               + `atacado_especializado`) / 3)


#####################################################################################################################################################################################################################################################

# PLANILHA 3 - INDICE AJUSTE SAZONAL COM VARIAcao TRIMESTRAL MOVEL

############## CALCULANDO AS VARIACOES TRIMESTRAIS ############

# CALCULO TRIMESTRE MOVEL VAREJISTA



# VARIACAO PERCENTUAL/VAREJISTA

variacao_trimestre_movel <- indice_ajuste_sazonal %>%
  
  mutate(variacao_tri_movel_varejista = ((media_movel_varejista / lag(media_movel_varejista,
                                                                      
                                                                      1)) - 1) * 100)



# CALCULO TRIMESTRE MOVEL AMPLIADO



# VARIACAO/PERCENTUAL AMPLIADO

variacao_trimestre_movel <- variacao_trimestre_movel %>%
  
  mutate(variacao_tri_movel_ampliado = ((media_movel_ampliado / lag(media_movel_ampliado,
                                                                    
                                                                    1)) - 1) * 100)




# CALCULO TRIMESTRE MOVEL COMBUSTIVEIS E LUBRIFICANTES




# VARIACAO/PERCENTUAL COMBLUB

variacao_trimestre_movel <- variacao_trimestre_movel %>%
  
  mutate(variacao_tri_movel_COMLUB = ((medial_movel_COMBLUB / lag(medial_movel_COMBLUB,
                                                                  
                                                                  1)) - 1) * 100)


# CALCULO TRIMESTRE MOVEL SUPERMERCADOS. HIPERMERCADOS, BEBIDAS E FUMO



# VARIACAO PERCENTUAL/ HSPBF

variacao_trimestre_movel <- variacao_trimestre_movel %>%
  mutate(variacao_tri_movel_HSPBF = ((media_movel_HSPBF / lag(media_movel_HSPBF,
                                                              1)) - 1) * 100)






# CALCULO TRIMESTRE MOVEL HIPERMERCADOS E SUPERMERCADOS


# VARIACAO PERCENTUAL/ HS

variacao_trimestre_movel <- variacao_trimestre_movel %>%
  mutate(variacao_tri_movel_HS = ((media_movel_HS / lag(media_movel_HS,
                                                        1)) - 1) * 100)




# CALCULO TRIMESTRE MOVEL TECIDOS,VESTUARIO E CALÇADOS

# VARIACAO PERCENTUAL TVC

variacao_trimestre_movel <- variacao_trimestre_movel %>%
  mutate(variacao_tri_movel_TVC = ((media_movel_TVC / lag(media_movel_TVC,
                                                          1)) - 1) * 100)




#  CALCULO TRIMESTRE MOVEL MOVEIS E ELETRODOMESTICOS

# VARIACAO PERCENTUAL ME

variacao_trimestre_movel <- variacao_trimestre_movel %>%
  mutate(variacao_tri_movel_ME = ((media_movel_ME / lag(media_movel_ME,
                                                        1)) - 1) * 100)




# CALCULO TRIMESTRE MOVEL MOVEIS

# VARIACAI PERCENTUAL MOVEIS

variacao_trimestre_movel <- variacao_trimestre_movel %>%
  mutate(variacao_tri_movel_moveis = ((media_movel_moveis / lag(media_movel_moveis,
                                                                1)) - 1) * 100)



# CALCULO TRIMESTRE MOVEL ELETRODOMESTICOS



# VARIACAO PERCENTUAL ELETRODOMESTICOS

variacao_trimestre_movel <- variacao_trimestre_movel %>%
  mutate(variacao_tri_movel_eletrodomesticos = ((media_movel_eletrodomesticos / lag(media_movel_eletrodomesticos,
                                                                                    1)) - 1) * 100)



# CALCULO TRIMESTRE MOVEL ARTIFGOS FARMACEUTICOS

# VARIACAO PERCEMNTUAL AMOPC

variacao_trimestre_movel <- variacao_trimestre_movel %>%
  mutate(variacao_tri_movel_AMOPC = ((media_movel_AMOPC / lag(media_movel_AMOPC,
                                                              1)) - 1) * 100)



# CALCULO TRIMESTRE MOVEL LIVROS,JORNAIS, REVISTAS E PAPELARA


# VARIACAO PERCENTUAL LJRP

variacao_trimestre_movel <- variacao_trimestre_movel %>%
  mutate(variacao_tri_movel_LJRP = ((media_movel_LJRP / lag(media_movel_LJRP,
                                                            1)) - 1) * 100)






# CALCULO TRIMESTRE MOVEL EQUIPAMENTOS E MATERIAIS PARA ESCRITORIO, INFORMATICA E COMUNICAÇÃO 


# VARIACAO PERCENTUAL EMEIC

variacao_trimestre_movel <- variacao_trimestre_movel %>%
  mutate(variacao_tri_movel_EMEIC = ((media_movel_EMEIC / lag(media_movel_EMEIC,
                                                              1)) - 1) * 100)




# CALCULO TRIMESTRE MOVELUSO PESSOAL DOMESTICO



# VARIACAO PERCENTUAL OAPD

variacao_trimestre_movel <- variacao_trimestre_movel %>%
  mutate(variacao_tri_movel_OAPD = ((media_movel_OAPD / lag(media_movel_OAPD,
                                                            1)) - 1) * 100)



# CALCULO TRIMESTRE MOVEL VEICULOS MOTOCICLETAS

# VARIACAO PERCENTUAL VMPP

variacao_trimestre_movel <- variacao_trimestre_movel %>%
  mutate(variacao_tri_movel_VMPP = ((media_movel_VMPP / lag(media_movel_VMPP,
                                                            1)) - 1) * 100)





# CALCULO TRIMESTRE MOVEL MATERIAL DE CONSTRUÇÃO


# VARIACAO PERCENTUAL MC

variacao_trimestre_movel<- variacao_trimestre_movel %>%
  mutate(variacao_tri_movel_MC = ((media_movel_MC / lag(media_movel_MC,
                                                        1)) - 1) * 100)






variacao_trimestre_movel<- variacao_trimestre_movel %>%
  mutate(variacao_tri_movel_AEPABF = ((media_movel_AEPABF / lag(media_movel_AEPABF,
                                                                1)) - 1) * 100)





# CALCULO TRIMESTRE MOVEL ATACADO ESPECIALIZADO


# VARIACAO PERCENTUAL AEPABF

variacao_trimestre_movel <- variacao_trimestre_movel %>%
  mutate(variacao_tri_movel_AEPABF = ((media_movel_AEPABF / lag(media_movel_AEPABF,
                                                                1)) - 1) * 100) %>% 
  
  select(Data,
         mes, 
         variacao_tri_movel_varejista, 
         variacao_tri_movel_ampliado, 
         variacao_tri_movel_COMLUB, 
         variacao_tri_movel_HSPBF, 
         variacao_tri_movel_HS,
         variacao_tri_movel_TVC,
         variacao_tri_movel_ME,
         variacao_tri_movel_moveis,
         variacao_tri_movel_eletrodomesticos,
         variacao_tri_movel_AMOPC,
         variacao_tri_movel_LJRP,
         variacao_tri_movel_EMEIC,
         variacao_tri_movel_OAPD,
         variacao_tri_movel_VMPP,
         variacao_tri_movel_MC,
         variacao_tri_movel_AEPABF)

###################################################################################################################################################################################################################################################

# PLANILHA 4 - INDICE AJUSTE SAZONAL COM VARIAÇÃO TRIMESTRAL  



# VARIACAO PERCENTUAL/VAREJISTA

variacao_trimestral <- indice_ajuste_sazonal %>%
  
  mutate(variacao_trimestral_varejista = ((media_movel_varejista / lag(media_movel_varejista,
                                                                       
                                                                       3)) - 1) * 100)




# VARIACAO/PERCENTUAL AMPLIADO

variacao_trimestral <- variacao_trimestral %>%
  
  mutate(variacao_trimestral_ampliado = ((media_movel_ampliado / lag(media_movel_ampliado,
                                                                     
                                                                     3)) - 1) * 100)




# VARIACAO/PERCENTUAL COMBLUB

variacao_trimestral <- variacao_trimestral %>%
  
  mutate(variacao_trimestral_COMLUB = ((medial_movel_COMBLUB / lag(medial_movel_COMBLUB,
                                                                   
                                                                   3)) - 1) * 100)





# VARIACAO PERCENTUAL/ HSPBF

variacao_trimestral <- variacao_trimestral %>%
  mutate(variacao_trimestral_HSPBF = ((media_movel_HSPBF / lag(media_movel_HSPBF,
                                                               3)) - 1) * 100)





# VARIACAO PERCENTUAL/ HS

variacao_trimestral <- variacao_trimestral %>%
  mutate(variacao_trimestral_HS = ((media_movel_HS / lag(media_movel_HS,
                                                         3)) - 1) * 100)




# VARIACAO PERCENTUAL TVC

variacao_trimestral <- variacao_trimestral %>%
  mutate(variacao_trimestral_TVC = ((media_movel_TVC / lag(media_movel_TVC,
                                                           3)) - 1) * 100)




# VARIACAO PERCENTUAL ME

variacao_trimestral <- variacao_trimestral %>%
  mutate(variacao_trimestral_ME = ((media_movel_ME / lag(media_movel_ME,
                                                         3)) - 1) * 100)





# VARIACAI PERCENTUAL MOVEIS

variacao_trimestral <- variacao_trimestral %>%
  mutate(variacao_trimestral_moveis = ((media_movel_moveis / lag(media_movel_moveis,
                                                                 3)) - 1) * 100)







# VARIACAO PERCENTUAL ELETRODOMESTICOS

variacao_trimestral <- variacao_trimestral %>%
  mutate(variacao_trimestral_eletrodomesticos = ((media_movel_eletrodomesticos / lag(media_movel_eletrodomesticos,
                                                                                     3)) - 1) * 100)







# VARIACAO PERCEMNTUAL AMOPC

variacao_trimestral <- variacao_trimestral %>%
  mutate(variacao_trimestrall_AMOPC = ((media_movel_AMOPC / lag(media_movel_AMOPC,
                                                                3)) - 1) * 100)





# VARIACAO PERCENTUAL LJRP

variacao_trimestral <- variacao_trimestral %>%
  mutate(variacao_trimestral_LJRP = ((media_movel_LJRP / lag(media_movel_LJRP,
                                                             3)) - 1) * 100)







# VARIACAO PERCENTUAL EMEIC

variacao_trimestral <- variacao_trimestral %>%
  mutate(variacao_trimestral_EMEIC = ((media_movel_EMEIC / lag(media_movel_EMEIC,
                                                               3)) - 1) * 100)









# VARIACAO PERCENTUAL OAPD

variacao_trimestral <- variacao_trimestral %>%
  mutate(variacao_trimestral_OAPD = ((media_movel_OAPD / lag(media_movel_OAPD,
                                                             3)) - 1) * 100)




# VARIACAO PERCENTUAL VMPP

variacao_trimestral <- variacao_trimestral %>%
  mutate(variacao_trimestral_VMPP = ((media_movel_VMPP / lag(media_movel_VMPP,
                                                             3)) - 1) * 100)




# VARIACAO PERCENTUAL MC

variacao_trimestral <- variacao_trimestral %>%
  mutate(variacao_trimestral_MC = ((media_movel_MC / lag(media_movel_MC,
                                                         3)) - 1) * 100)






# VARIACAO PERCENTUAL AEPABF

variacao_trimestral <- variacao_trimestral %>%
  mutate(variacao_trimestral_AEPABF = ((media_movel_AEPABF / lag(media_movel_AEPABF,
                                                                 3)) - 1) * 100) %>%  
  
  select(Data, 
         mes, 
         variacao_trimestral_varejista, 
         variacao_trimestral_ampliado, 
         variacao_trimestral_COMLUB, 
         variacao_trimestral_HSPBF, 
         variacao_trimestral_HS,
         variacao_trimestral_TVC,
         variacao_trimestral_ME,
         variacao_trimestral_moveis,
         variacao_trimestral_eletrodomesticos,
         variacao_trimestrall_AMOPC,
         variacao_trimestral_LJRP,
         variacao_trimestral_EMEIC,
         variacao_trimestral_OAPD,
         variacao_trimestral_VMPP,
         variacao_trimestral_MC,
         variacao_trimestral_AEPABF)




############################################################################################################################################################################################################################################




# PLANILHA 5 - VARIAÇÃO MES A MES IMEDIANTAMENTE ANTERIOR COM AJUSTE SAZONAL


#            EXTRAINDO E TRATANDO DADOS TABELAS 80/81/83 ANTERIOR COM AJUSTE SAZONAL


variacao_mes_ant_ajuste_sazonal = get_sidra(api = "/t/8880/n1/all/v/11708/p/all/c11046/56734/d/v11708%201")%>%
  select(Valor, `Mês (Código)`)
#mutate(Valor = as.numeric(Valor))
names(variacao_mes_ant_ajuste_sazonal)[1] <- "variacao_mes_ant_ajuste_sazonal"


variacao_mes_ant_ajuste_sazonal_ampliado = get_sidra(api = "/t/8881/n1/all/v/11708/p/all/c11046/56736/d/v11708%201")%>%
  select( `Mês (Código)`, Valor)
#mutate(Valor = as.numeric(Valor))
names(variacao_mes_ant_ajuste_sazonal_ampliado)[2] <- "variacao_mes_ant_ajuste_sazonal_ampliado"




variacao_mes_ant_ajuste_sazonal_ampliado_atividades = get_sidra(api = "/t/8883/n1/all/v/11708/p/all/c11046/56736/c85/all/d/v11708%201")%>%
  select(Valor, Atividades, `Mês (Código)`)
#mutate(Valor = as.numeric(Valor))




# PIVOTANDO VARIACAO_MES_ANT_AJUSTE_SAZONAL E GERANDO UM NOVO DATAFRAME

df3 <- pivot_wider(variacao_mes_ant_ajuste_sazonal_ampliado_atividades, names_from = Atividades, values_from = c("Valor"))


# CRIANDO DATAFRAME CC

cc <- variacao_mes_ant_ajuste_sazonal[,c("Mês (Código)",
                                         "variacao_mes_ant_ajuste_sazonal")]

# MERGE DOS DATAFRAMES

variacao_mes_anterior_ajuste_sazonal <- merge(cc,variacao_mes_ant_ajuste_sazonal_ampliado, by = "Mês (Código)")

# MERGE DO MERGE

variacao_mes_anterior_ajuste_sazonal <- merge(variacao_mes_anterior_ajuste_sazonal,df3, by = "Mês (Código)")









# ADICIONANDO COLUNA DATA EM VARIACAO_MES_ANTERIOR_AJUSTE_SAZONAL


variacao_mes_anterior_ajuste_sazonal <- variacao_mes_anterior_ajuste_sazonal %>% dplyr::mutate(Data = as.Date(paste0(substr(variacao_mes_anterior_ajuste_sazonal$`Mês (Código)`, 
                                                                                                                            start = 1, stop = 4),"-",
                                                                                                                     substr(variacao_mes_anterior_ajuste_sazonal$`Mês (Código)`, 
                                                                                                                            start = 5 , stop = 6), "-01")))


# ALOCANDO DATA NA PRIMEIRA COLUNA


variacao_mes_anterior_ajuste_sazonal <- variacao_mes_anterior_ajuste_sazonal %>%
  select(Data, everything())

colnames(variacao_mes_anterior_ajuste_sazonal) <- c("Data",
                                                    "mes",
                                                    "var_mensal_varejista",
                                                    "var_mensal_ampliado",
                                                    "combustíveis_e_lubrificantes",
                                                    "hipermer_supermer_prod_alimen_bebidas",
                                                    "hipermercados_supermercados",
                                                    "tecidos_vestuário_calçados",
                                                    "moveis_eletrodomesticos",
                                                    "moveis",
                                                    "eletrodomesticos",
                                                    "artigos_farmacêuticos_medicos_cosmeticos",
                                                    "livros_jornais_papelaria",
                                                    "equipamentos_materiais_escritorio_informatica",
                                                    "outros_artigos",
                                                    "veiculos_motocicletas",
                                                    "material_de_construcao",
                                                    "atacado_especializado")


#######################################################################################################################################################################################################################################################



#PLANILHA 6 - VARIAÇÃO INTERANUAL

# EXTRAINDO DADOS TABELA 80/81/83 

variacao_mes_ano_anterior_varejista = get_sidra(api = "/t/8880/n1/all/v/11709/p/all/c11046/56734/d/v11709%201")%>%
  select(`Mês (Código)`, Mês, Valor)
#mutate(Valor = as.numeric(Valor))
names(variacao_mes_ano_anterior_varejista)[3] <- "variacao_mes_ano_anterior"
#View(variacao_mes_ano_anterior_varejista)


variacao_mes_ano_anterior_ampliado = get_sidra(api = "/t/8881/n1/all/v/11709/p/all/c11046/56736/d/v11709%201")%>%
  select(Valor, `Mês (Código)`)
#mutate(Valor = as.numeric(Valor))
names(variacao_mes_ano_anterior_ampliado)[1] <- "variacao_mes_ano_anterior_ampliado"



variacao_mes_ano_anterior_ampliado_atividades = get_sidra(api = "/t/8883/n1/all/v/11709/p/all/c11046/56736/c85/all/d/v11709%201")%>%
  select(Valor, `Mês (Código)`, Atividades)
#mutate(Valor = as.numeric(Valor))




# PIVOTANDO VARIACAO_MES_ANO_ANTERIOR_AMPLIADO_ATIVIDADES E GERANDO NOVO DATAFRAME


df4 <- pivot_wider(variacao_mes_ano_anterior_ampliado_atividades, names_from = Atividades, values_from = c("Valor"))


# CRIANDO DATAFRAME DD

dd <- variacao_mes_ano_anterior_varejista[,c("Mês (Código)",
                                             "variacao_mes_ano_anterior")]
# MERGE

variacao_mes_ano_anterior <- merge(dd,variacao_mes_ano_anterior_ampliado, by = "Mês (Código)")


# MERGE DO MERGE

variacao_mes_ano_anterior <- merge(variacao_mes_ano_anterior, df4, by = "Mês (Código)")

#CRIANDO O DF VARIAÇÃO INTERANUAL

variacao_interanual <- variacao_mes_ano_anterior

# CRIANDO COLUNA DATA EM VARIACAO_MES_ANTERIOR

variacao_interanual <- variacao_interanual %>% dplyr::mutate(Data = as.Date(paste0(substr(variacao_interanual$`Mês (Código)`, 
                                                                                          start = 1, stop = 4),"-",
                                                                                   substr(variacao_interanual$`Mês (Código)`, 
                                                                                          start = 5 , stop = 6), "-01")))
# ALOCANDO DATA NA PRIMEIRA COLUNA

variacao_interanual <- variacao_interanual %>%
  select(Data, everything())


#### RENOMEANDO AS COLUNA DO DATAFRAME

colnames(variacao_interanual) <- c("Data",
                                   "mes", 
                                   "var_mes_ano_anterior_varejista",
                                   "var_mes_ano_anterior_ampliado",
                                   "combustíveis_e_lubrificantes",
                                   "hipermer_supermer_prod_alimen_bebidas",
                                   "hipermercados_supermercados",
                                   "tecidos_vestuário_calçados",
                                   "moveis_eletrodomesticos",
                                   "moveis",
                                   "eletrodomesticos",
                                   "artigos_farmacêuticos_medicos_cosmeticos",
                                   "livros_jornais_papelaria",
                                   "equipamentos_materiais_escritorio_informatica",
                                   "outros_artigos",
                                   "veiculos_motocicletas",
                                   "material_de_construcao",
                                   "atacado_especializado")





######################################################################################################################################################################################################################################



# PLANILHA 7 - VARIAÇÃO ACOMULADA NO ANO (EM RELAÇÃO AO MESMO PERIODO DO ANOS ANTERIOR %)

# EXTRAINDO DADOS E TRATANDO VARIACAO ACUMULADA NO ANO



variacao_acumulada_ano_varejista = get_sidra(api = "/t/8880/n1/all/v/11710/p/all/c11046/56734/d/v11710%201")%>%
  select(`Mês (Código)`, Valor)
names(variacao_acumulada_ano_varejista)[2] <- "variacao_acumulada_ano_varejista"



variacao_acumulada_ano_ampliado = get_sidra(api = "/t/8881/n1/all/v/11710/p/all/c11046/56736/d/v11710%201")%>%
  select(`Mês (Código)`, Valor)
names(variacao_acumulada_ano_ampliado)[2] <- "variacao_acumulada_ano_ampliado"




variacao_acumulada_ano_ampliado_atividades = get_sidra(api = "/t/8883/n1/all/v/11710/p/all/c11046/56736/c85/all/d/v11710%201")%>%
  select(`Mês (Código)`, Valor, Atividades)


# PIVOTANDO VARIACAO_ACUMULADA_ANO_ATIVIDADES E GERANDO UM NOVO DATAFRAME

df5 <- pivot_wider(variacao_acumulada_ano_ampliado_atividades, names_from = Atividades, values_from = c("Valor"))



# CRIANDO DATAFRAME EE

ee <- variacao_acumulada_ano_varejista [,c("Mês (Código)",
                                           "variacao_acumulada_ano_varejista")]


# MERGE

variacao_acumulada_ano <- merge(ee, variacao_acumulada_ano_ampliado, by = "Mês (Código)")



# MERGE DO MERGE

variacao_acumulada_ano <- merge(variacao_acumulada_ano, df5, by = "Mês (Código)")


# ADICIONANDO A COLUNA DATA


variacao_acumulada_ano <- variacao_acumulada_ano %>% dplyr::mutate(Data = as.Date(paste0(substr(variacao_acumulada_ano$`Mês (Código)`, 
                                                                                                start = 1, stop = 4),"-",
                                                                                         substr(variacao_acumulada_ano$`Mês (Código)`,
                                                                                                start = 5 , stop = 6), "-01")))
# ALOCANDO DATA NA PRIMEIRA COLUNA

variacao_acumulada_ano <- variacao_acumulada_ano %>%
  select(Data, everything())

#selecionando as variaveis do dataframe Variaca_acomulada_ano



colnames(variacao_acumulada_ano) <- c("Data",
                                      "mes",
                                      "variacao_acumulada_ano_varejista",
                                      "variacao_acumulada_ano_ampliado",
                                      "variacao_acumulada_ano_combustíveis_e_lubrificantes",
                                      "variacao_acumulada_ano_hipermer_supermer_prod_alimen_bebidas",
                                      "variacao_acumulada_ano_hipermercados_supermercados",
                                      "variacao_acumulada_ano_tecidos_vestuário_calçados",
                                      "variacao_acumulada_ano_moveis_eletrodomesticos",
                                      "variacao_acumulada_ano_moveis",
                                      "variacao_acumulada_ano_eletrodomesticos",
                                      "variacao_acumulada_ano_artigos_farmacêuticos_medicos_cosmeticos",
                                      "variacao_acumulada_ano_livros_jornais_papelaria",
                                      "variacao_acumulada_ano_equipamentos_materiais_escritorio_informatica",
                                      "variacao_acumulada_ano_outros_artigos",
                                      "variacao_acumulada_ano_veiculos_motocicletas",
                                      "variacao_acumulada_ano_material_de_construcao",
                                      "variacao_acumulada_ano_atacado_especializado")


###################################################################################################################################################################################################################################################


# PLANILHA 8 - VARIAÇÃO ACOMULADA EM 12 MESES 

# EXTRAINDO DADOS E TRATANDO VARIACAO ACUMULADA EM 12 MESES


variacao_acumulada_12meses_varejista = get_sidra(api= "/t/8880/n1/all/v/11711/p/all/c11046/56734/d/v11711%201")%>%
  select(Mês, `Mês (Código)`, Valor)
names(variacao_acumulada_12meses_varejista)[3] <- "variacao_acumulada_12meses_varejista"



variacao_acumulada_12meses_ampliado = get_sidra(api= "/t/8881/n1/all/v/11711/p/all/c11046/56736/d/v11711%201")%>%
  select(`Mês (Código)`, Valor)
names(variacao_acumulada_12meses_ampliado)[2] <- "variacao_acumulada_12meses_ampliado"



variacao_acumulada_12meses_ampliado_atividades = get_sidra(api = "/t/8883/n1/all/v/11711/p/all/c11046/56736/c85/all/d/v11711%201")%>%
  select(`Mês (Código)`, Valor, Atividades)





# PIVOTANDO VARIACAO_ACUMULADA_12MESES_AMPLIADO_ATIVIDADES E GERANDO UM NOVO DATAFRAME

df6 <- pivot_wider(variacao_acumulada_12meses_ampliado_atividades, names_from = Atividades, values_from = c("Valor"))


# CRIANDO DATAFRAME FF

ff <- variacao_acumulada_12meses_varejista[,c("Mês (Código)",
                                              "variacao_acumulada_12meses_varejista")]

# MERGE


variacao_acumulada_12meses <- merge(ff, variacao_acumulada_12meses_ampliado, by = "Mês (Código)")


# MERGE DO MERGE

variacao_acumulada_12meses <- merge(variacao_acumulada_12meses, df6, by = "Mês (Código)")





# CRIANDO COLUNA DATA EM VARIACAO_ACUMULADA12MESES

variacao_acumulada_12meses <- variacao_acumulada_12meses %>% dplyr::mutate(Data = as.Date(paste0(substr(variacao_acumulada_12meses$`Mês (Código)`, 
                                                                                                        start = 1, stop = 4),"-",
                                                                                                 substr(variacao_acumulada_12meses$`Mês (Código)`,
                                                                                                        start = 5 , stop = 6), "-01")))

#  ALOCANDO DATA NA PRIMEIRA COLUNA

variacao_acumulada_12meses <- variacao_acumulada_12meses %>%
  select(Data, everything())


colnames(variacao_acumulada_12meses) <- c("Data",
                                          "mes",
                                          "var_12m_varejista",
                                          "var_12_ampliado",
                                          "combustíveis_e_lubrificantes",
                                          "hipermer_supermer_prod_alimen_bebidas",
                                          "hipermercados_supermercados",
                                          "tecidos_vestuário_calçados",
                                          "moveis_eletrodomesticos",
                                          "moveis",
                                          "eletrodomesticos",
                                          "artigos_farmacêuticos_medicos_cosmeticos",
                                          "livros_jornais_papelaria",
                                          "equipamentos_materiais_escritorio_informatica",
                                          "outros_artigos",
                                          "veiculos_motocicletas",
                                          "material_de_construcao",
                                          "atacado_especializado")



#ESPECIFCANDO O CAMINHO PARA O ARQUIVO XLSX

caminho_arquivo <- "C://Users//Usuario//Desktop//PESQUISA M COMERCIO//PMC_BASE.xlsx"


# Crie um arquivo XLSX

wb <- createWorkbook()


# Adicione a primeira planilha


addWorksheet(wb, sheetName = "Numero_indice")


writeData(wb, sheet = "Numero_indice", x = numero_indice)


# Adicione a segunda planilha

addWorksheet(wb, sheetName = "Indice_ajuste_sazonal")



writeData(wb, sheet = "Indice_ajuste_sazonal", x = numero_indice_ajuste_sazonal)





# Adicione a terceira planilha



addWorksheet(wb, sheetName = "Variacao_trimestral_movel")



writeData(wb, sheet = "Variacao_trimestral_movel", x = variacao_trimestre_movel)





# Adicione a quarta planilha



addWorksheet(wb, sheetName = "Variacao_trimestral")



writeData(wb, sheet = "Variacao_trimestral", x = variacao_trimestral)


# Adicione a quinta planilha



addWorksheet(wb, sheetName = "Var_mensal_com_ajuste")



writeData(wb, sheet = "Var_mensal_com_ajuste", x = variacao_mes_anterior_ajuste_sazonal)





# Adicione a sexta planilha



addWorksheet(wb, sheetName = "variacao_interanual")



writeData(wb, sheet = "variacao_interanual", x = variacao_interanual)




# Adicione a setima planilha



addWorksheet(wb, sheetName = "var_acomulada_ano")



writeData(wb, sheet = "var_acomulada_ano", x = variacao_acumulada_ano)


# Adicione a oitava planilha



addWorksheet(wb, sheetName = "var_acomulada_12meses")



writeData(wb, sheet = "var_acomulada_12meses", x = variacao_acumulada_12meses)


# Salve o arquivo XLSX



saveWorkbook(wb, caminho_arquivo, overwrite = TRUE)


