import argparse
import getpass
from CRUD import criar_tabela, criar_usuario, SETOR_SAC, SETORES_VALIDOS

def main():
    parser = argparse.ArgumentParser(description="Cria um usuário de login do sistema.")
    parser.add_argument("--usuario", help="Nome de usuário")
    parser.add_argument("--senha", help="Senha (evite deixar visível em histórico compartilhado)")
    parser.add_argument("--setor", choices=SETORES_VALIDOS, help="Setor do usuário (SAC ou Qualidade)")
    args = parser.parse_args()

    criar_tabela()

    usuario = args.usuario or input("Nome de usuário: ").strip()

    if args.senha:
        senha = args.senha
    else:
        senha = getpass.getpass("Senha: ")
        confirmacao = getpass.getpass("Confirme a senha: ")
        if senha != confirmacao:
            print("As senhas não coincidem.")
            return

    setor = args.setor
    if not setor:
        escolha = input(f"Setor ({'/'.join(SETORES_VALIDOS)}) [{SETOR_SAC}]: ").strip()
        setor = escolha or SETOR_SAC
        if setor not in SETORES_VALIDOS:
            print(f"Setor inválido. Use um dos seguintes: {', '.join(SETORES_VALIDOS)}.")
            return

    if not usuario or not senha:
        print("Usuário e senha são obrigatórios.")
        return

    sucesso, erro = criar_usuario(usuario, senha, setor)

    if sucesso:
        print(f'Usuário "{usuario}" criado com sucesso no setor "{setor}".')
    else:
        print(f"Erro ao criar usuário: {erro}")

if __name__ == "__main__":
    main()
