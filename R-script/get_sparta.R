
#if(require("readxl"))  install.packages("readxl")
#if(require("stringr")) install.packages("stringr")
#if(require("xlsx"))    install.packages("xlsx")

library("stringr")
library("readxl")
library("xlsx")

# Limpar memória
rm(list = ls())

# Diretório dados SPARTA
SPARTA_DIR = '../SPARTA/'

# Listar arquivos SPARTA no diretório (Somente concessionárias - Excluindo Roraima)
lista_arqs = list.files(path=SPARTA_DIR)
lista_arqs = lista_arqs[!str_detect(lista_arqs,pattern="~")]
lista_arqs = lista_arqs[!str_detect(lista_arqs,pattern="Roraima")]
lista_arqs = lista_arqs[str_detect(lista_arqs,pattern="SPARTA")]


# Inicializando vetores para armazenamento dos dados das concessionárias
Concessionaria = c()

UC_faixa_1 = c()
UC_faixa_2 = c()
UC_faixa_3 = c()
UC_faixa_4 = c()

Consumo_MWh_faixa_1 = c()
Consumo_MWh_faixa_2 = c()
Consumo_MWh_faixa_3 = c()
Consumo_MWh_faixa_4 = c()

# Adquirindo dados de interesse por concessionária
for(nome_arq in lista_arqs)
{
  print(nome_arq)

  dados_SPARTA = read_excel(paste(SPARTA_DIR,nome_arq,sep = ""), sheet="Mercado_Receita")
  
  Concessionaria = c(Concessionaria, nome_arq)

  # Residencial baixa renda - faixa 01
  dados_SPARTA_faixa_1 = subset(dados_SPARTA,dados_SPARTA$Subclasse=="Residencial baixa renda – faixa 01")
  dados_SPARTA_faixa_1 = subset(dados_SPARTA_faixa_1,dados_SPARTA_faixa_1$TipoMercado=="Regular")
  Consumo_MWh_faixa_1 = c(Consumo_MWh_faixa_1, 12*mean(dados_SPARTA_faixa_1$TUSD_E))
  UC_faixa_1          = c(UC_faixa_1, mean(dados_SPARTA_faixa_1$UC))

  # Residencial baixa renda - faixa 02
  dados_SPARTA_faixa_2 = subset(dados_SPARTA,dados_SPARTA$Subclasse=="Residencial baixa renda – faixa 02")
  dados_SPARTA_faixa_2 = subset(dados_SPARTA_faixa_2,dados_SPARTA_faixa_2$TipoMercado=="Regular")
  Consumo_MWh_faixa_2 = c(Consumo_MWh_faixa_2, 12*mean(dados_SPARTA_faixa_2$TUSD_E))
  UC_faixa_2          = c(UC_faixa_2, mean(dados_SPARTA_faixa_2$UC))

  # Residencial baixa renda - faixa 03
  dados_SPARTA_faixa_3 = subset(dados_SPARTA,dados_SPARTA$Subclasse=="Residencial baixa renda – faixa 03")
  dados_SPARTA_faixa_3 = subset(dados_SPARTA_faixa_3,dados_SPARTA_faixa_3$TipoMercado=="Regular")
  Consumo_MWh_faixa_3 = c(Consumo_MWh_faixa_3, 12*mean(dados_SPARTA_faixa_3$TUSD_E))
  UC_faixa_3          = c(UC_faixa_3, mean(dados_SPARTA_faixa_3$UC))
  
  # Residencial baixa renda - faixa 04
  dados_SPARTA_faixa_4 = subset(dados_SPARTA,dados_SPARTA$Subclasse=="Residencial baixa renda – faixa 04")
  dados_SPARTA_faixa_4 = subset(dados_SPARTA_faixa_4,dados_SPARTA_faixa_4$TipoMercado=="Regular")
  Consumo_MWh_faixa_4 = c(Consumo_MWh_faixa_4, 12*mean(dados_SPARTA_faixa_4$TUSD_E))
  UC_faixa_4          = c(UC_faixa_4, mean(dados_SPARTA_faixa_4$UC))
}

MWp_faixa_1 = rep(-1,length(lista_arqs))
MWp_faixa_2 = rep(-1,length(lista_arqs))
MWp_faixa_3 = rep(-1,length(lista_arqs))
MWp_faixa_4 = rep(-1,length(lista_arqs))

MW_faixa_1 = rep(-1,length(lista_arqs))
MW_faixa_2 = rep(-1,length(lista_arqs))
MW_faixa_3 = rep(-1,length(lista_arqs))
MW_faixa_4 = rep(-1,length(lista_arqs))

df = data.frame(Concessionaria, Consumo_MWh_faixa_1,UC_faixa_1, MWp_faixa_1, MW_faixa_1, MWp_faixa_1, MW_faixa_1,
                                Consumo_MWh_faixa_2,UC_faixa_2, MWp_faixa_2, MW_faixa_2, MWp_faixa_1, MW_faixa_1,
                                Consumo_MWh_faixa_3,UC_faixa_3, MWp_faixa_3, MW_faixa_3, MWp_faixa_1, MW_faixa_1,
                                Consumo_MWh_faixa_4,UC_faixa_4, MWp_faixa_4, MW_faixa_4, MWp_faixa_1, MW_faixa_1)

write.xlsx2(df, paste0("SPARTA_TODOS.xlsx"), row.names = FALSE)
