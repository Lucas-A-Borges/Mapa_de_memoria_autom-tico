
# Gerador Automático de Mapa de Memória

Programa para geração do **mapa de memória** de forma automática a partir das programações realizadas no **Control Expert**.

---

## 📄 Informações do Projeto

- **Criado por:** Lucas Alves Borges  
- **Versão do Programa:** 2  
- **Data:** 09/12/2025  
- **Versão Control Expert:** 16
- **Versão PLC:** M580 Schneider  
- **Versão Python:** 3.13.9  

---

## 📘 Instruções de Uso

1. Exporte o arquivo **ZEF** do PLC.  
2. Abra o arquivo exportado com uma ferramenta de descompactação (**WinRAR** ou **7zip**).  
3. Extraia o arquivo **`unitpro.xef`**.  
4. Coloque na mesma pasta os seguintes arquivos:
   - `unitpro.xef`  
   - `modelo_mapa_memoria.xlsx`  
   - `Gerar_mapa_de_memoria.exe`
5. Execute o arquivo **`Gerar_mapa_de_memoria.exe`**.

---

## ⚠️ Considerações

- Os arquivos **`unitpro.xef`** e **`modelo_mapa_memoria.xlsx`** **não podem** estar abertos durante a execução.
- O programa pode levar algum tempo para rodar e concluir o processo.
- Podem ocorrer lacunas no mapa de memória devido à falta de padronização total dos programas.

---

## ❗ Possíveis Erros e Soluções

- Caso o arquivo não abra, tente executá-lo diretamente pelo **Prompt de Comando**.
- Se nenhum equipamento for gerado no mapa de memória, tente **extrair novamente** o arquivo `unitpro.xef`.

---

## 🧩 PLC's Já Padronizados

- IT1000CN01  
- IT1000CN07  
- IT1470CN01  
