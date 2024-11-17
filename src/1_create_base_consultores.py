import random
from datetime import datetime, timedelta

import holidays
import polars as pl
from faker import Faker
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo


def obter_feriados(ano):
    feriados = holidays.Brazil(years=[ano])
    feriados_uteis = [data for data in feriados.keys() if data.weekday() < 5]
    return set(feriados_uteis)


def gerar_faltas_aleatorias(ano, consultores, num_faltas=3):
    faltas = []
    inicio_ano = datetime(ano, 1, 1)
    fim_ano = datetime(ano, 12, 31)
    dias_uteis = []
    data_atual = inicio_ano
    while data_atual <= fim_ano:
        if data_atual.weekday() < 5:
            dias_uteis.append(data_atual.strftime("%Y-%m-%d"))
        data_atual += timedelta(days=1)

    for consultor in consultores:
        faltas_consultor = random.sample(dias_uteis, num_faltas)
        for falta in faltas_consultor:
            faltas.append(
                {
                    "Consultor_ID": consultor["ID"],
                    "Nome": consultor["Nome"],
                    "Data_Falta": falta,
                    "Horas": random.randint(2, 8),
                }
            )

    return faltas


def gerar_faltas_feriados(ano, consultores, feriados):
    faltas_feriados = []
    for consultor in consultores:
        for feriado in feriados:
            faltas_feriados.append(
                {
                    "Consultor_ID": consultor["ID"],
                    "Nome": consultor["Nome"],
                    "Data_Falta": feriado.strftime("%Y-%m-%d"),
                    "Horas": 8,
                }
            )
    return faltas_feriados


def main():
    try:
        faker = Faker("pt_BR")
        consultores = [
            {"ID": i + 1, "Nome": faker.unique.name()} for i in range(5)
        ]

        feriados_2023 = obter_feriados(2023)
        feriados_2024 = obter_feriados(2024)

        faltas_2023_aleatorias = gerar_faltas_aleatorias(2023, consultores)
        faltas_2024_aleatorias = gerar_faltas_aleatorias(2024, consultores)

        faltas_2023_feriados = gerar_faltas_feriados(
            2023, consultores, feriados_2023
        )
        faltas_2024_feriados = gerar_faltas_feriados(
            2024, consultores, feriados_2024
        )

        faltas = (
            faltas_2023_aleatorias
            + faltas_2024_aleatorias
            + faltas_2023_feriados
            + faltas_2024_feriados
        )

        df_faltas = pl.DataFrame(faltas)

        caminho_arquivo = "src/bases_excel/faltas_consultores.xlsx"
        df_faltas.write_excel(caminho_arquivo)

        print(
            f"Base de faltas salva e formatada como tabela com sucesso em: {caminho_arquivo}"
        )

    except Exception as e:
        print(f"Erro ao gerar a base de faltas: {str(e)}")


if __name__ == "__main__":
    main()
