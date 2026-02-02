from filtre_service import filter_ocorrencias

print("ANTES DA FUNÇÃO")

df, path = filter_ocorrencias()

print("DEPOIS DA FUNÇÃO")
print(df.head())
print("Arquivo gerado em:", path)
