import os
import random
from datetime import datetime, timedelta

import polars as pl


# Função para criar dados de backlog
def gerar_backlogs():
    equipes = ["Teste Analytics", "Dev. Analytics", "Análise de Requisitos"]
    consultores = [
        "Ana Nogueira",
        "Enzo Farias",
        "Ísis Monteiro",
        "Pietro Cassiano",
        "Bárbara Andrade",
    ]
    projetos = ["Projeto 1", "Projeto 2", "Projeto 3"]
    chamados = ["Projeto1", "Projeto2", "Projeto3"]
    titulos = [
        "Laboratório de dados 1",
        "Laboratório de dados 2",
        "Laboratório de dados 3",
    ]
    prioridades = ["1 - Alta", "2 - Média", "3 - Baixa", "4 - Sem Prioridade"]

    data_inicio = datetime(2023, 1, 2)
    sprints = []

    # Gera datas de sprints de 15 dias
    while data_inicio.year <= 2024:
        data_fim = data_inicio + timedelta(days=14)
        sprints.append((data_inicio, data_fim))
        data_inicio = data_fim + timedelta(days=1)

    backlogs = []
    id_tarefa = 4200

    for sprint_id, (inicio, fim) in enumerate(sprints, start=1):
        for _ in range(random.randint(1, 3)):  # Até 3 backlogs por sprint
            equipe = random.choice(equipes)
            consultor = random.choice(consultores)
            projeto = random.choice(projetos)
            chamado = random.choice(chamados)
            titulo = random.choice(titulos)
            prioridade = random.choice(prioridades)
            valor_negocio = random.randint(20, 200)
            effort = random.randint(8, 40)

            # Define a data de fechamento com 75% de probabilidade de ser na data prevista
            fechou_na_data = random.random() < 0.75
            data_fechamento = (
                fim
                if fechou_na_data
                else fim + timedelta(days=random.randint(15, 60))
            )

            backlog = {
                "Area Path": equipe,
                "Assigned To": consultor,
                "Business Value": valor_negocio,
                "Changed Date": data_fechamento.strftime("%Y-%m-%d"),
                "Closed Date": data_fechamento.strftime("%Y-%m-%d"),
                "Created Date": inicio.strftime("%Y-%m-%d"),
                "Cycle Time Days": (data_fechamento - inicio).days,
                "Data Fim Previsto": fim.strftime("%Y-%m-%d"),
                "Data Início Previsto": inicio.strftime("%Y-%m-%d"),
                "Effort": effort,
                "Iteration End Date": fim.strftime("%Y-%m-%d"),
                "Iteration Path": f"{equipe}/Sprint {sprint_id}",
                "Iteration Start Date": inicio.strftime("%Y-%m-%d"),
                "Lead Time Days": (data_fechamento - inicio).days,
                "Priority": prioridade,
                "Project Name": projeto,
                "Service Now": chamado,
                "Tags": "",
                "Title": titulo,
                "Value Area": "Business",
                "Versão": "1.0",
                "Work Item Id": id_tarefa,
                "Work Item Type": "Product Backlog Item",
            }

            backlogs.append(backlog)
            id_tarefa += 1

    return backlogs


# Função para criar tarefas a partir dos backlogs
def gerar_tasks(backlogs):
    atividades = [
        "arquitetura de dados",
        "documentação",
        "engenharia de dados",
        "criação de painéis",
    ]
    tasks = []

    for backlog in backlogs:
        sprint_inicio = datetime.strptime(
            backlog["Data Início Previsto"], "%Y-%m-%d"
        )
        sprint_fim = datetime.strptime(
            backlog["Data Fim Previsto"], "%Y-%m-%d"
        )

        num_tasks = random.randint(1, 3)  # Até 3 tarefas por backlog
        for i in range(num_tasks):
            atividade = random.choice(atividades)
            assigned_to = backlog["Assigned To"]
            iteration_path = backlog["Iteration Path"]
            estado = "Concluído" if random.random() < 0.8 else "Não Concluído"
            fechado_dentro_sprint = (
                sprint_fim
                if estado == "Concluído"
                else sprint_fim + timedelta(days=random.randint(1, 15))
            )

            task = {
                "Activity": atividade,
                "Area Path": backlog["Area Path"],
                "Assigned To": assigned_to,
                "Changed Date": fechado_dentro_sprint.strftime("%Y-%m-%d"),
                "Closed Date": fechado_dentro_sprint.strftime("%Y-%m-%d")
                if estado == "Concluído"
                else "",
                "Created Date": sprint_inicio.strftime("%Y-%m-%d"),
                "Iteration Path": iteration_path,
                "Parent Work Item Id": backlog["Work Item Id"],
                "State": estado,
                "Title": atividade,
                "Work Item Id": f"{backlog['Work Item Id']}-{i + 1}",  # ID de task baseado no backlog
                "Work Item Type": "Task",
            }
            tasks.append(task)

    return tasks


# Função para salvar os dados no Excel
def salvar_excel(backlogs, tasks, arquivo_backlog, arquivo_task):
    try:
        os.makedirs(os.path.dirname(arquivo_backlog), exist_ok=True)

        # Criar DataFrame com Polars para backlogs e tasks
        df_backlog = pl.DataFrame(backlogs)
        df_task = pl.DataFrame(tasks)

        # Salvar como Excel
        df_backlog.write_excel(arquivo_backlog)
        df_task.write_excel(arquivo_task)

        print(
            f"Arquivos Excel salvos com sucesso: {arquivo_backlog}, {arquivo_task}"
        )

    except Exception as e:
        print(f"Erro ao salvar os arquivos Excel: {e}")


# Código principal
if __name__ == "__main__":
    try:
        backlogs = gerar_backlogs()
        tasks = gerar_tasks(backlogs)
        salvar_excel(
            backlogs,
            tasks,
            "src/bases_excel/product_backlog.xlsx",
            "src/bases_excel/tasks.xlsx",
        )
    except Exception as e:
        print(f"Erro ao gerar os dados: {e}")