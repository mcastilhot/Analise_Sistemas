# Inventário de Hardware e Diagnóstico de Sistema — Topocart Consultoria

Aplicação desktop em Python (PySide6) para coleta automatizada de especificações de hardware, diagnóstico de gargalos de desempenho e geração consolidada de relatórios em formatos **Word (.docx)** e **Excel (.xlsx)**.

Compatível com **Windows** e **Linux**.

---

## 🚀 Funcionalidades

- **Coleta Completa de Hardware & Sistema**:
  - **Processador (CPU)**: Modelo, núcleos, threads e status de virtualização na BIOS (VT-x / AMD-V).
  - **Memória RAM**: Capacidade total, frequência configurada (MHz), geração (DDR4 / DDR5), canais (Single vs Dual Channel) e detalhamento por slot.
  - **Placas de Vídeo (GPU) & Telas**: Identificação de GPUs integradas e dedicadas, medição precisa de VRAM (NVIDIA via `nvidia-smi`), resoluções e taxas de atualização (Hz).
  - **Armazenamento**: Detecção de discos físicos (NVMe PCIe vs SATA SSD vs HDD mecânico), status de saúde e monitoramento de espaço livre em todas as partições montadas.
  - **Rede & Conectividade**: Tipo de interface (Ethernet vs Wi-Fi), velocidade do link (1 Gbps, 10 Gbps, 100 Mbps), IP local, Gateway e MAC Address.
  - **Sistema & Energia**: Edição/Build do Windows, tempo de atividade contínua (Uptime) e plano de energia ativo.
  - **Placa-mãe & BIOS**: Fabricante, modelo da placa-mãe, versão e data de lançamento da BIOS.

- **Diagnóstico Automático de Gargalos**:
  - Alerta de RAM abaixo de 32 GB e operação em Single Channel.
  - Alerta de falta de GPU dedicada para fluxos pesados.
  - Alerta de cabos/portas de rede limitadas a 100 Mbps em vez de Gigabit.
  - Alerta de pouco espaço livre no disco C: e unidades secundárias.
  - Alerta de ausência de discos NVMe/SSD e planos de energia em economia.

- **Relatórios Consolidados**:
  - **.docx**: Relatório formatado em orientação paisagem com tabela executiva, diagnóstico por máquina e fichas técnicas individuais detalhadas (pronto para Word e Google Docs).
  - **.xlsx**: Planilha estruturada com colunas formatadas, cabeçalhos estilizados e congelamento de painéis (pronto para Excel e Google Sheets).

---

## 📦 Instalação

Clone o repositório ou baixe os arquivos:

```bash
git clone https://github.com/mcastilhot/inventario-hardware.git
cd inventario-hardware
```

Instale as dependências com Python 3.10+:

```bash
pip install -r requirements.txt
```

---

## 🖥️ Como Usar

Execute o aplicativo:

```bash
python inventario_hardware.py
```

1. **Preencha os campos da máquina**:
   - Apelido (ex: `Workstation-01`)
   - Responsável (preenchido automaticamente com o usuário do sistema)
   - Função principal (ex: `Processamento LiDAR`, `CAD`, `Administrativo`)
   - Observações (opcional)
2. **Selecione a Pasta de Inventário**:
   - Para inventário em rede, selecione uma pasta compartilhada acessível por todas as máquinas.
3. **Clique em "▶ Analisar e Salvar Esta Máquina"**:
   - Os dados são coletados e salvos em formato JSON individual dentro da pasta.
4. **Clique em "Gerar Relatório Consolidado"**:
   - O sistema compila todos os arquivos JSON encontrados na pasta e gera automaticamente os relatórios `.docx` e `.xlsx`.
