from mate_pipeline import main as run_pipeline

if __name__ == "__main__":
    valor = input("Data/URL/caminho: ").strip() or None
    run_pipeline(valor)
