# Disponibilidade Comercial

Aplicativo em Python para a equipe comercial consultar produtos disponíveis e ainda sem venda, separados por cultura e produto. A atualização fica em uma tela protegida por senha; a consulta fica liberada para os consultores.

## O que a primeira versão faz

- Consulta por cultura, produto, origem e observação.
- Filtro para ocultar itens sem estoque.
- Cadastro e atualização de disponibilidade por cultura/produto/origem.
- Status: `Disponivel`, `Baixo estoque` e `Sem estoque`.
- Visual simples, responsivo e fácil de usar no Android ou iPhone pelo navegador.
- Banco local SQLite em `data/disponibilidade.db`.
- Dependência principal: Streamlit.

## Rodar no computador

```bash
pip install -r requirements.txt
streamlit run app.py
```

Depois acesse o endereço mostrado pelo Streamlit, normalmente:

```text
http://localhost:8501
```

## Senha de atualização

Para testar localmente, a senha padrão é:

```text
admin123
```

Antes de publicar, altere a senha usando uma variável de ambiente ou o arquivo de segredos do Streamlit:

```toml
ADMIN_PASSWORD = "uma-senha-forte"
```

## Hospedagem gratuita

A forma mais simples é publicar no Streamlit Community Cloud:

1. Crie uma conta gratuita em https://streamlit.io/cloud.
2. Suba estes arquivos para um repositório no GitHub.
3. No Streamlit Cloud, clique em `New app`.
4. Selecione o repositório e informe `app.py` como arquivo principal.
5. Em `Settings > Secrets`, cadastre:

```toml
ADMIN_PASSWORD = "sua-senha-aqui"
```

6. Compartilhe o link publicado com os consultores.

## Observação importante sobre dados

Esta primeira versão usa SQLite porque é simples e rápida para começar. Em hospedagens gratuitas, o arquivo do banco pode ser perdido em reinícios ou redeploys. Para uso diário com vários usuários, a próxima evolução recomendada é trocar o SQLite por:

- Google Sheets, para uma operação simples e fácil de auditar.
- Supabase, para um banco gratuito mais robusto com login e histórico.

## Próximos incrementos recomendados

- Login individual para cada atualizador.
- Histórico de alterações.
- Exportação para Excel.
- Campo de data prevista de venda ou vencimento.
- Notificação automática quando um produto ficar disponível.
