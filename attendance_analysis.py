import pandas as pd


def analyze_meeting_attendance(file_path):
    # Carrega a aba "2.Participants"
    df = pd.read_excel(file_path, sheet_name='2.Participants')

    # Exibe os nomes das colunas para confirmar os nomes exatos
    print("Colunas disponíveis no DataFrame:", df.columns)


# Caminho para o arquivo Excel
file_path = '/Users/jaquelinesilva/Library/CloudStorage/Dropbox/My Mac (MacBook-Pro.local)/Downloads/DELETE DOWNLOADS/untitled folder/Ferramentas da qualidade (5159.A1) - Attendance report 10-28-24 (2).xlsx'

# Chamar a função
analyze_meeting_attendance(file_path)
import pandas as pd

def analyze_meeting_attendance(file_path):
    # Tenta carregar a aba "2.Participants" e pula as primeiras linhas até os dados principais
    df = pd.read_excel(file_path, sheet_name='2.Participants', skiprows=3)  # Ajuste o número de linhas conforme necessário

    # Exibe os nomes das colunas para confirmar se agora estão corretos
    print("Colunas disponíveis no DataFrame:", df.columns)

    # Continue apenas se o DataFrame carregar corretamente
    if 'Name' in df.columns:
        # Defina o nome das colunas para corresponder ao arquivo
        nome_coluna_participante = 'Name'
        coluna_duracao = 'In-Meeting Duration'

        # Remove duplicatas e calcula o tempo total
        df_unique = df.drop_duplicates(subset=[nome_coluna_participante])
        df_unique[coluna_duracao] = pd.to_timedelta(df_unique[coluna_duracao])

        # Calcula o tempo total de todos os participantes
        total_duracao = df_unique[coluna_duracao].sum()
        numero_participantes = df_unique[nome_coluna_participante].nunique()

        print(f"Número de participantes únicos: {numero_participantes}")
        print(f"Tempo total de participação na reunião: {total_duracao}")
    else:
        print("Erro: Coluna 'Name' não encontrada. Verifique o arquivo.")

# Caminho para o arquivo Excel
file_path = '/Users/jaquelinesilva/Library/CloudStorage/Dropbox/My Mac (MacBook-Pro.local)/Downloads/DELETE DOWNLOADS/untitled folder/Ferramentas da qualidade (5159.A1) - Attendance report 10-28-24 (2).xlsx'

# Chamar a função
analyze_meeting_attendance(file_path)
import pandas as pd


def analyze_meeting_attendance(file_path):
    # Define os nomes das colunas que sabemos existir no arquivo
    colunas = ['Name', 'First Join', 'Last Leave', 'In-Meeting Duration', 'Email', 'User ID', 'Role']

    # Carrega a aba "2.Participants" e aplica os nomes das colunas
    df = pd.read_excel(file_path, sheet_name='2.Participants', skiprows=3, names=colunas)

    # Exibe os nomes das colunas para verificar se estão corretos
    print("Colunas disponíveis no DataFrame:", df.columns)

    # Remover duplicatas para participantes únicos
    df_unique = df.drop_duplicates(subset=['Name'])

    # Converte a duração para timedelta e calcula o total
    df_unique['In-Meeting Duration'] = pd.to_timedelta(df_unique['In-Meeting Duration'], errors='coerce')

    # Calcula o tempo total de todos os participantes
    total_duracao = df_unique['In-Meeting Duration'].sum()
    numero_participantes = df_unique['Name'].nunique()

    print(f"Número de participantes únicos: {numero_participantes}")
    print(f"Tempo total de participação na reunião: {total_duracao}")


# Caminho para o arquivo Excel
file_path = '/Users/jaquelinesilva/Library/CloudStorage/Dropbox/My Mac (MacBook-Pro.local)/Downloads/DELETE DOWNLOADS/untitled folder/Ferramentas da qualidade (5159.A1) - Attendance report 10-28-24 (2).xlsx'

# Chamar a função
analyze_meeting_attendance(file_path)
import pandas as pd


def analyze_meeting_attendance(file_path):
    # Define os nomes das colunas para o arquivo
    colunas = ['Name', 'First Join', 'Last Leave', 'In-Meeting Duration', 'Email', 'User ID', 'Role']

    # Carrega a aba "2.Participants" e aplica os nomes das colunas
    df = pd.read_excel(file_path, sheet_name='2.Participants', skiprows=3, names=colunas)

    # Exibe os nomes das colunas para verificar se estão corretos
    print("Colunas disponíveis no DataFrame:", df.columns)

    # Remove duplicatas para participantes únicos
    df_unique = df.drop_duplicates(subset=['Name']).copy()

    # Converte a duração para timedelta e calcula o total
    df_unique.loc[:, 'In-Meeting Duration'] = pd.to_timedelta(df_unique['In-Meeting Duration'], errors='coerce')

    # Calcula o tempo total de todos os participantes
    total_duracao = df_unique['In-Meeting Duration'].sum()
    numero_participantes = df_unique['Name'].nunique()

    print(f"Número de participantes únicos: {numero_participantes}")
    print(f"Tempo total de participação na reunião: {total_duracao}")


# Caminho para o arquivo Excel
file_path = '/Users/jaquelinesilva/Library/CloudStorage/Dropbox/My Mac (MacBook-Pro.local)/Downloads/DELETE DOWNLOADS/untitled folder/Ferramentas da qualidade (5159.A1) - Attendance report 10-28-24 (2).xlsx'

# Chamar a função
analyze_meeting_attendance(file_path)
import pandas as pd


def analyze_meeting_attendance(file_path):
    # Define os nomes das colunas para o arquivo
    colunas = ['Name', 'First Join', 'Last Leave', 'In-Meeting Duration', 'Email', 'User ID', 'Role']

    # Carrega a aba "2.Participants" e aplica os nomes das colunas
    df = pd.read_excel(file_path, sheet_name='2.Participants', skiprows=3, names=colunas)

    # Exibe os nomes das colunas para verificar se estão corretos
    print("Colunas disponíveis no DataFrame:", df.columns)

    # Remove duplicatas para participantes únicos e converte a duração para timedelta
    df_unique = df.drop_duplicates(subset=['Name']).assign(
        **{'In-Meeting Duration': pd.to_timedelta(df['In-Meeting Duration'], errors='coerce')}
    )

    # Calcula o tempo total de todos os participantes
    total_duracao = df_unique['In-Meeting Duration'].sum()
    numero_participantes = df_unique['Name'].nunique()

    print(f"Número de participantes únicos: {numero_participantes}")
    print(f"Tempo total de participação na reunião: {total_duracao}")


# Caminho para o arquivo Excel
file_path = '/Users/jaquelinesilva/Library/CloudStorage/Dropbox/My Mac (MacBook-Pro.local)/Downloads/DELETE DOWNLOADS/untitled folder/Ferramentas da qualidade (5159.A1) - Attendance report 10-28-24 (2).xlsx'

# Chamar a função
analyze_meeting_attendance(file_path)
import pandas as pd


def analyze_meeting_attendance(file_path):
    # Define os nomes das colunas para o arquivo
    colunas = ['Name', 'First Join', 'Last Leave', 'In-Meeting Duration', 'Email', 'User ID', 'Role']

    # Carrega a aba "2.Participants" e aplica os nomes das colunas
    df = pd.read_excel(file_path, sheet_name='2.Participants', skiprows=3, names=colunas)

    # Remove duplicatas para participantes únicos e converte a duração para timedelta
    df['In-Meeting Duration'] = pd.to_timedelta(df['In-Meeting Duration'], errors='coerce')

    # Agrupa por nome e soma a duração para cada participante
    duracao_por_participante = df.groupby('Name')['In-Meeting Duration'].sum()

    # Exibe o número de participantes únicos
    numero_participantes = duracao_por_participante.count()
    print(f"Número de participantes únicos: {numero_participantes}")

    # Exibe o tempo total de participação para cada participante
    print("\nTempo total de participação de cada participante:")
    for participante, duracao in duracao_por_participante.iteritems():
        print(f"{participante}: {duracao}")


# Caminho para o arquivo Excel
file_path = '/Users/jaquelinesilva/Library/CloudStorage/Dropbox/My Mac (MacBook-Pro.local)/Downloads/DELETE DOWNLOADS/untitled folder/Ferramentas da qualidade (5159.A1) - Attendance report 10-28-24 (2).xlsx'

# Chamar a função
analyze_meeting_attendance(file_path)
import pandas as pd


def analyze_meeting_attendance(file_path):
    # Define os nomes das colunas para o arquivo
    colunas = ['Name', 'First Join', 'Last Leave', 'In-Meeting Duration', 'Email', 'User ID', 'Role']

    # Carrega a aba "2.Participants" e aplica os nomes das colunas
    df = pd.read_excel(file_path, sheet_name='2.Participants', skiprows=3, names=colunas)

    # Remove duplicatas para participantes únicos e converte a duração para timedelta
    df['In-Meeting Duration'] = pd.to_timedelta(df['In-Meeting Duration'], errors='coerce')

    # Agrupa por nome e soma a duração para cada participante
    duracao_por_participante = df.groupby('Name')['In-Meeting Duration'].sum()

    # Exibe o número de participantes únicos
    numero_participantes = duracao_por_participante.count()
    print(f"Número de participantes únicos: {numero_participantes}")

    # Exibe o tempo total de participação para cada participante
    print("\nTempo total de participação de cada participante:")
    for participante, duracao in duracao_por_participante.items():
        print(f"{participante}: {duracao}")


# Caminho para o arquivo Excel
file_path = '/Users/jaquelinesilva/Library/CloudStorage/Dropbox/My Mac (MacBook-Pro.local)/Downloads/DELETE DOWNLOADS/untitled folder/Ferramentas da qualidade (5159.A1) - Attendance report 10-28-24 (2).xlsx'

# Chamar a função
analyze_meeting_attendance(file_path)
import pandas as pd


def analyze_meeting_attendance(file_path):
    # Define os nomes das colunas para o arquivo
    colunas = ['Name', 'First Join', 'Last Leave', 'In-Meeting Duration', 'Email', 'User ID', 'Role']

    # Carrega a aba "2.Participants" e aplica os nomes das colunas
    df = pd.read_excel(file_path, sheet_name='2.Participants', skiprows=3, names=colunas)

    # Remove duplicatas para participantes únicos e converte a duração para timedelta
    df['In-Meeting Duration'] = pd.to_timedelta(df['In-Meeting Duration'], errors='coerce')

    # Agrupa por nome e soma a duração para cada participante
    duracao_por_participante = df.groupby('Name')['In-Meeting Duration'].sum()

    # Exibe o número de participantes únicos
    numero_participantes = duracao_por_participante.count()
    print(f"Número de participantes únicos: {numero_participantes}")

    # Exibe o tempo total de participação para cada participante
    print("\nTempo total de participação de cada participante:")
    for participante, duracao in duracao_por_participante.items():  # Correção aqui com .items()
        print(f"{participante}: {duracao}")


# Caminho para o arquivo Excel
file_path = '/Users/jaquelinesilva/Library/CloudStorage/Dropbox/My Mac (MacBook-Pro.local)/Downloads/DELETE DOWNLOADS/untitled folder/Ferramentas da qualidade (5159.A1) - Attendance report 10-28-24 (2).xlsx'

# Chamar a função
analyze_meeting_attendance(file_path)
import pandas as pd


def analyze_meeting_attendance(file_path):
    # Define os nomes das colunas para o arquivo
    colunas = ['Name', 'First Join', 'Last Leave', 'In-Meeting Duration', 'Email', 'User ID', 'Role']

    # Carrega a aba "2.Participants" e aplica os nomes das colunas
    df = pd.read_excel(file_path, sheet_name='2.Participants', skiprows=3, names=colunas)

    # Remove duplicatas para participantes únicos e converte a duração para timedelta
    df['In-Meeting Duration'] = pd.to_timedelta(df['In-Meeting Duration'], errors='coerce')

    # Agrupa por nome e soma a duração para cada participante
    duracao_por_participante = df.groupby('Name')['In-Meeting Duration'].sum()

    # Exibe o número de participantes únicos
    numero_participantes = duracao_por_participante.count()
    print(f"Número de participantes únicos: {numero_participantes}")

    # Exibe o tempo total de participação para cada participante
    print("\nTempo total de participação de cada participante:")
    for participante, duracao in duracao_por_participante.items():  # Certifique-se que .items() está aqui
        print(f"{participante}: {duracao}")


# Caminho para o arquivo Excel
file_path = '/Users/jaquelinesilva/Library/CloudStorage/Dropbox/My Mac (MacBook-Pro.local)/Downloads/DELETE DOWNLOADS/untitled folder/Ferramentas da qualidade (5159.A1) - Attendance report 10-28-24 (2).xlsx'

# Chamar a função
analyze_meeting_attendance(file_path)
import pandas as pd


def analyze_meeting_attendance(file_path):
    # Define os nomes das colunas para o arquivo
    colunas = ['Name', 'First Join', 'Last Leave', 'In-Meeting Duration', 'Email', 'User ID', 'Role']

    # Carrega a aba "2.Participants" e aplica os nomes das colunas
    df = pd.read_excel(file_path, sheet_name='2.Participants', skiprows=3, names=colunas)

    # Remove duplicatas para participantes únicos e converte a duração para timedelta
    df['In-Meeting Duration'] = pd.to_timedelta(df['In-Meeting Duration'], errors='coerce')

    # Agrupa por nome e soma a duração para cada participante
    duracao_por_participante = df.groupby('Name')['In-Meeting Duration'].sum()

    # Exibe o número de participantes únicos
    numero_participantes = duracao_por_participante.count()
    print(f"Número de participantes únicos: {numero_participantes}")

    # Exibe o tempo total de participação para cada participante
    print("\nTempo total de participação de cada participante:")
    for participante, duracao in duracao_por_participante.items():  # Certifique-se que .items() está aqui
        print(f"{participante}: {duracao}")


# Caminho para o arquivo Excel
file_path = '/Users/jaquelinesilva/Library/CloudStorage/Dropbox/My Mac (MacBook-Pro.local)/Downloads/DELETE DOWNLOADS/untitled folder/Ferramentas da qualidade (5159.A1) - Attendance report 10-28-24 (2).xlsx'

# Chamar a função
analyze_meeting_attendance(file_path)


