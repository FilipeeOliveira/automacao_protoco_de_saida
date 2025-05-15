import re

class CpfUtils:
    @staticmethod
    def formatar_cpf(cpf: str) -> str:
        """Formata o CPF para o padrão XXX.XXX.XXX-XX, somente se tiver 11 dígitos"""
        cpf_numeros = CpfUtils.somente_numeros(cpf)
        if len(cpf_numeros) != 11:
            return cpf  # Retorna o original se não tiver 11 dígitos
        return f"{cpf_numeros[:3]}.{cpf_numeros[3:6]}.{cpf_numeros[6:9]}-{cpf_numeros[9:]}"

    @staticmethod
    def somente_numeros(cpf: str) -> str:
        """Remove qualquer caractere não numérico"""
        return re.sub(r'\D', '', cpf)

    @staticmethod
    def validar_cpf(cpf: str) -> bool:
        """Valida o CPF com base no algoritmo oficial"""
        cpf = CpfUtils.somente_numeros(cpf)
        if len(cpf) != 11 or cpf == cpf[0] * 11:
            return False

        def calc_digito(cpf_parcial):
            soma = sum(int(digito) * peso for digito, peso in zip(cpf_parcial, range(len(cpf_parcial) + 1, 1, -1)))
            resto = soma % 11
            return '0' if resto < 2 else str(11 - resto)

        digito1 = calc_digito(cpf[:9])
        digito2 = calc_digito(cpf[:9] + digito1)
        return cpf[-2:] == digito1 + digito2

    @staticmethod
    def limitar_11_digitos(valor: str) -> str:
        """Limita o CPF a 11 números"""
        numeros = CpfUtils.somente_numeros(valor)
        return numeros[:11]
