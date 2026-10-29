#!/usr/bin/env python3
"""Inventário de Hardware e Diagnóstico de Sistema - Topocart Consultoria (Windows / Linux)

Uso: Preencha apelido/responsável/função -> Analisar (salva um JSON na pasta de
inventário). Em qualquer máquina, "Gerar relatório" junta todos os JSONs da pasta
em um .docx (editável no Google Docs / Word) e um .xlsx (Excel / Sheets).

Dependências: pip install PySide6 python-docx openpyxl
"""
from __future__ import annotations

import datetime
import getpass
import json
import re
import shutil
import socket
import subprocess
import sys
from pathlib import Path

# ============================================================== COLETA
NO_WINDOW = 0x08000000 if sys.platform == "win32" else 0


def run(cmd: list[str] | str, timeout: int = 40) -> str:
    try:
        return subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout,
            creationflags=NO_WINDOW,
        ).stdout.strip()
    except Exception:
        return ""


def ps_json(cmd: str) -> list[dict]:
    """Executa um comando no PowerShell e converte a saída para lista de dicionários."""
    out = run(["powershell", "-NoProfile", "-Command", f"{cmd} | ConvertTo-Json -Compress"])
    if not out:
        return []
    try:
        d = json.loads(out)
        return d if isinstance(d, list) else [d]
    except Exception:
        return []


def nvidia() -> list[dict]:
    """Coleta dados precisos de VRAM e driver usando nvidia-smi."""
    out = run(["nvidia-smi", "--query-gpu=name,memory.total,driver_version", "--format=csv,noheader"])
    gpus = []
    for linha in out.splitlines():
        p = [x.strip() for x in linha.split(",")]
        if len(p) == 3:
            gpus.append({"nome": p[0], "vram": p[1], "driver": p[2]})
    return gpus


def _smbios_mem_type(tipo: int | None) -> str:
    mapa = {
        20: "DDR", 21: "DDR2", 22: "DDR2 FB-DIMM", 24: "DDR3",
        26: "DDR4", 29: "LPDDR4", 30: "DDR5", 31: "LPDDR5", 34: "DDR5", 35: "LPDDR5"
    }
    return mapa.get(tipo or 0, "DDR" if (tipo or 0) > 0 else "")


def coletar_windows() -> dict:
    g = lambda c: (ps_json(c) or [{}])[0]

    # CPU e Virtualização
    cpu = g("Get-CimInstance Win32_Processor")
    cs = g("Get-CimInstance Win32_ComputerSystem")
    os_info = g("Get-CimInstance Win32_OperatingSystem")
    bios = g("Get-CimInstance Win32_BIOS")
    mb = g("Get-CimInstance Win32_BaseBoard")

    # Memória RAM detalhada
    mems = ps_json("Get-CimInstance Win32_PhysicalMemory")
    ram_modulos = []
    bancos_usados = set()
    speeds = []
    tipos_ram = set()
    for m in mems:
        cap_gb = round((m.get("Capacity") or 0) / (1024**3), 1)
        speed = m.get("ConfiguredClockSpeed") or m.get("Speed") or 0
        bank = m.get("BankLabel") or m.get("DeviceLocator") or "?"
        tipo_str = _smbios_mem_type(m.get("SMBIOSMemoryType"))
        if tipo_str:
            tipos_ram.add(tipo_str)
        if speed:
            speeds.append(speed)
        if bank != "?":
            bancos_usados.add(bank)
        ram_modulos.append({
            "slot": m.get("DeviceLocator", "?"),
            "banco": bank,
            "capacidade_gb": cap_gb,
            "velocidade_mhz": speed,
            "tipo": tipo_str,
            "fabricante": m.get("Manufacturer", "").strip(),
            "part_number": (m.get("PartNumber") or "").strip(),
        })

    total_ram_gb = round((cs.get("TotalPhysicalMemory") or 0) / (1024**3))
    canal = "Dual Channel" if len(bancos_usados) >= 2 or (len(ram_modulos) >= 2 and len(ram_modulos) % 2 == 0) else "Single Channel"
    ram_tipo = "/".join(sorted(tipos_ram)) if tipos_ram else "DDR"
    ram_mhz = max(speeds) if speeds else 0

    # Placa de Vídeo e Resolução
    video_controllers = ps_json("Get-CimInstance Win32_VideoController")
    gpus = []
    telas = []
    for vc in video_controllers:
        nome_gpu = vc.get("Name", "")
        if not re.search("Remote|Basic|Virtual", nome_gpu, re.I):
            gpus.append(nome_gpu)
        res_h = vc.get("CurrentHorizontalResolution")
        res_v = vc.get("CurrentVerticalResolution")
        hz = vc.get("CurrentRefreshRate")
        if res_h and res_v:
            telas.append(f"{res_h}x{res_v} @ {hz or 60}Hz ({nome_gpu})")

    # Armazenamento: Discos Físicos e Barramentos
    discos_fisicos = []
    for d in ps_json("Get-PhysicalDisk"):
        nome_d = d.get("FriendlyName", "?")
        tipo_midia = d.get("MediaType", "?")
        bus = d.get("BusType", "")
        tamanho_gb = round((d.get("Size") or 0) / 1e9)
        saude = d.get("HealthStatus", "Healthy")
        bus_str = f" [{bus}]" if bus and bus not in ("?", "Unknown") else ""
        discos_fisicos.append({
            "nome": nome_d,
            "tipo": tipo_midia,
            "bus": bus,
            "tamanho_gb": tamanho_gb,
            "saude": saude,
            "texto": f"{nome_d} ({tipo_midia}{bus_str}, {tamanho_gb} GB, Saúde: {saude})",
        })

    # Armazenamento: Volumes Lógicos / Partições montadas
    volumes = []
    for v in ps_json("Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3'"):
        letra = v.get("DeviceID", "")
        tot = round((v.get("Size") or 0) / 1e9)
        livre = round((v.get("FreeSpace") or 0) / 1e9)
        pct_livre = round((livre / tot * 100), 1) if tot > 0 else 0
        volumes.append({
            "drive": letra,
            "nome": v.get("VolumeName", ""),
            "sistema_arquivos": v.get("FileSystem", ""),
            "total_gb": tot,
            "livre_gb": livre,
            "pct_livre": pct_livre,
            "texto": f"{letra} {livre} GB livres de {tot} GB ({pct_livre}%)",
        })

    # Disco C principal
    disco_c = next((vol for vol in volumes if vol["drive"].upper() == "C:"), None)
    uso_c_total = disco_c["total_gb"] if disco_c else 0
    uso_c_livre = disco_c["livre_gb"] if disco_c else 0

    # Rede: Adaptadores Ativos e Link Speed
    net_adapters = ps_json("Get-NetAdapter | Where-Object Status -eq 'Up'")
    net_configs = ps_json("Get-CimInstance Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=True'")
    redes = []
    for na in net_adapters:
        nome_int = na.get("Name", "")
        desc = na.get("InterfaceDescription", "")
        speed = na.get("LinkSpeed", "")
        mac = na.get("MacAddress", "")
        tipo_fisico = na.get("PhysicalMediaType") or na.get("MediaType") or ""
        is_wifi = "802.11" in str(tipo_fisico) or "Wi-Fi" in nome_int or "Wireless" in desc

        cfg = next((c for c in net_configs if c.get("MACAddress") == mac or c.get("Description") == desc), {})
        ips = cfg.get("IPAddress", [])
        if isinstance(ips, str):
            ips = [ips]
        ipv4 = next((ip for ip in ips if "." in ip), "")
        gateway = (cfg.get("DefaultIPGateway") or [""])[0] if isinstance(cfg.get("DefaultIPGateway"), list) else (cfg.get("DefaultIPGateway") or "")
        dns_list = cfg.get("DNSServerSearchOrder") or []
        dns_str = ", ".join(dns_list) if isinstance(dns_list, list) else str(dns_list)

        redes.append({
            "nome": nome_int,
            "descricao": desc,
            "tipo": "Wi-Fi" if is_wifi else "Ethernet",
            "link_speed": speed,
            "mac": mac,
            "ipv4": ipv4,
            "gateway": gateway,
            "dns": dns_str,
            "texto": f"{nome_int} ({'Wi-Fi' if is_wifi else 'Ethernet'} {speed}): IP {ipv4 or '—'}",
        })

    # Sistema Operacional, Uptime e Plano de Energia
    uptime_dias = 0
    last_boot = os_info.get("LastBootUpTime")
    if last_boot:
        match = re.search(r"\d+", str(last_boot))
        if match:
            boot_ts = int(match.group(0)) / 1000.0
            uptime_dias = round((datetime.datetime.now().timestamp() - boot_ts) / 86400, 1)

    power_out = run("powercfg /getactivescheme")
    plano_energia = "Equilibrado"
    if "Alto desempenho" in power_out or "High performance" in power_out:
        plano_energia = "Alto Desempenho"
    elif "Economia de energia" in power_out or "Power saver" in power_out:
        plano_energia = "Economia de Energia"
    elif "Desempenho máximo" in power_out or "Ultimate Performance" in power_out:
        plano_energia = "Desempenho Máximo"
    elif "(" in power_out and ")" in power_out:
        plano_energia = power_out.split("(")[-1].split(")")[0].strip()

    data_bios = ""
    bios_date_raw = bios.get("ReleaseDate", "")
    match_bd = re.search(r"\d+", str(bios_date_raw))
    if match_bd:
        data_bios = datetime.datetime.fromtimestamp(int(match_bd.group(0)) / 1000.0).strftime("%d/%m/%Y")

    return {
        "fabricante": cs.get("Manufacturer"),
        "modelo": cs.get("Model"),
        "serial": bios.get("SerialNumber"),
        "placa_mae": f"{mb.get('Manufacturer', '')} {mb.get('Product', '')}".strip(),
        "bios_versao": bios.get("SMBIOSBIOSVersion"),
        "bios_data": data_bios,
        "cpu": (cpu.get("Name") or "").strip(),
        "nucleos": cpu.get("NumberOfCores"),
        "threads": cpu.get("NumberOfLogicalProcessors"),
        "virtualizacao_habilitada": bool(cpu.get("VirtualizationFirmwareEnabled")),
        "ram_gb": total_ram_gb,
        "ram_mhz": ram_mhz,
        "ram_tipo": ram_tipo,
        "ram_canal": canal,
        "ram_modulos": ram_modulos,
        "ram_slots_usados": len(ram_modulos),
        "gpu": gpus,
        "telas": telas,
        "discos": [d["texto"] for d in discos_fisicos],
        "discos_detalhe": discos_fisicos,
        "volumes": volumes,
        "disco_sistema_total_gb": uso_c_total,
        "disco_sistema_livre_gb": uso_c_livre,
        "rede": redes,
        "so": f"{os_info.get('Caption', '')} (Build {os_info.get('BuildNumber', '')}, {os_info.get('OSArchitecture', '')})".strip(),
        "uptime_dias": uptime_dias,
        "plano_energia": plano_energia,
    }


def coletar_linux() -> dict:
    lscpu = {}
    for line in run(["lscpu"]).splitlines():
        if ":" in line:
            k, v = line.split(":", 1)
            lscpu[k.strip()] = v.strip()

    ram_kb = 0
    try:
        for line in Path("/proc/meminfo").read_text().splitlines():
            if line.startswith("MemTotal"):
                ram_kb = int(line.split()[1])
    except Exception:
        pass

    def dmi(n: str) -> str:
        try:
            return Path(f"/sys/class/dmi/id/{n}").read_text().strip()
        except Exception:
            return ""

    gpus = [line.split(": ", 1)[-1] for line in run(["lspci"]).splitlines() if "VGA" in line or "3D controller" in line]

    discos_fisicos = []
    for line in run(["lsblk", "-dno", "NAME,MODEL,SIZE,ROTA", "-e", "7,11"]).splitlines():
        p = line.split()
        if len(p) >= 4:
            tipo = "HDD" if p[-1] == "1" else "SSD/NVMe"
            discos_fisicos.append({
                "nome": " ".join(p[1:-2]),
                "tipo": tipo,
                "texto": f"{' '.join(p[1:-2])} ({tipo}, {p[-2]})",
            })

    volumes = []
    for line in run(["df", "-hP"]).splitlines()[1:]:
        parts = line.split()
        if len(parts) >= 6 and parts[0].startswith("/dev/"):
            volumes.append({
                "drive": parts[5],
                "total_gb": parts[1],
                "livre_gb": parts[3],
                "texto": f"{parts[5]} ({parts[3]} livres de {parts[1]})",
            })

    so = ""
    try:
        for line in Path("/etc/os-release").read_text().splitlines():
            if line.startswith("PRETTY_NAME"):
                so = line.split("=", 1)[1].strip('"')
    except Exception:
        pass

    socks = int(lscpu.get("Socket(s)", 1) or 1)
    uso = shutil.disk_usage("/")
    uptime_str = run(["uptime", "-p"])

    return {
        "fabricante": dmi("sys_vendor"),
        "modelo": dmi("product_name"),
        "serial": dmi("product_serial"),
        "placa_mae": f"{dmi('board_vendor')} {dmi('board_name')}".strip(),
        "bios_versao": dmi("bios_version"),
        "bios_data": dmi("bios_date"),
        "cpu": lscpu.get("Model name", ""),
        "nucleos": int(lscpu.get("Core(s) per socket", 0) or 0) * socks,
        "threads": int(lscpu.get("CPU(s)", 0) or 0),
        "virtualizacao_habilitada": "VT-x" in lscpu.get("Virtualization", "") or "AMD-V" in lscpu.get("Virtualization", ""),
        "ram_gb": round(ram_kb / (1024**2)),
        "ram_mhz": 0,
        "ram_tipo": "DDR",
        "ram_canal": "Desconhecido",
        "ram_modulos": [],
        "ram_slots_usados": 0,
        "gpu": gpus,
        "telas": [],
        "discos": [d["texto"] for d in discos_fisicos],
        "discos_detalhe": discos_fisicos,
        "volumes": volumes,
        "disco_sistema_total_gb": round(uso.total / 1e9),
        "disco_sistema_livre_gb": round(uso.free / 1e9),
        "rede": [],
        "so": so,
        "uptime_dias": 0,
        "plano_energia": uptime_str,
    }


def coletar(cadastro: dict) -> dict:
    d = coletar_windows() if sys.platform == "win32" else coletar_linux()
    d["gpu_nvidia"] = nvidia()
    d["hostname"] = socket.gethostname()
    d["usuario_logado"] = getpass.getuser()
    d["data_coleta"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    d["cadastro"] = cadastro
    return d


# ============================================================== ANÁLISE E GARGALOS
DEDICADA = re.compile(r"NVIDIA|GeForce|Quadro|RTX|GTX|Radeon RX|Radeon Pro|Arc ", re.I)


def gpu_texto(m: dict) -> str:
    if m.get("gpu_nvidia"):
        return " | ".join(f"{g['nome']} ({g['vram']})" for g in m["gpu_nvidia"])
    return " | ".join(m.get("gpu") or []) or "—"


def tem_gpu_dedicada(m: dict) -> bool:
    return bool(m.get("gpu_nvidia")) or any(DEDICADA.search(g) for g in m.get("gpu") or [])


def gargalos(m: dict) -> list[str]:
    """Detecta de forma inteligente gargalos de hardware, conectividade e configuração."""
    g = []

    # Memória RAM
    ram = m.get("ram_gb") or 0
    if ram < 16:
        g.append(f"RAM crítica de apenas {ram} GB (mínimo recomendado para operações técnicas: 32 GB)")
    elif ram < 32:
        g.append(f"RAM de {ram} GB (abaixo de 32 GB recomendados para fluxos pesados)")

    if m.get("ram_canal") == "Single Channel":
        g.append("RAM operando em Single Channel (largura de banda de memória reduzida pela metade; recomendado instalar pentes em pares para Dual Channel)")

    # Processador e Virtualização
    threads = m.get("threads") or 0
    if threads < 8:
        g.append(f"CPU limitada com apenas {threads} threads")

    if not m.get("virtualizacao_habilitada", True):
        g.append("Virtualização (VT-x / AMD-V) desativada na BIOS (bloqueia WSL2, Docker e máquinas virtuais)")

    # Placa de Vídeo
    if not tem_gpu_dedicada(m):
        g.append("Sem GPU dedicada (utilizando apenas vídeo integrado)")

    # Armazenamento e Espaço
    livre_c = m.get("disco_sistema_livre_gb") or 0
    if livre_c < 50:
        g.append(f"Espaço crítico no disco C: ({livre_c} GB livres)")
    elif livre_c < 100:
        g.append(f"Pouco espaço livre no disco C: ({livre_c} GB)")

    # Discos secundários cheios
    for vol in m.get("volumes") or []:
        if isinstance(vol, dict) and vol.get("drive", "").upper() != "C:":
            if (vol.get("livre_gb") or 0) < 30 or (vol.get("pct_livre") or 100) < 8:
                g.append(f"Partição {vol.get('drive')} com pouco espaço livre ({vol.get('livre_gb')} GB, {vol.get('pct_livre')}%)")

    discos_str = " ".join(m.get("discos") or [])
    if discos_str and not re.search(r"SSD|NVMe", discos_str, re.I):
        g.append("Sem SSD/NVMe no sistema (apenas discos mecânicos HDD)")
    elif discos_str and not re.search(r"NVMe", discos_str, re.I):
        g.append("Sem disco NVMe (utilizando apenas SATA; recomendado NVMe para arquivos volumosos)")

    # Conectividade e Rede
    for r in m.get("rede") or []:
        speed = str(r.get("link_speed", ""))
        tipo = r.get("tipo", "")
        if "100 Mbps" in speed and tipo == "Ethernet":
            g.append(f"Rede cabeada negociando a apenas 100 Mbps na interface '{r.get('nome')}' (cabo de rede ou porta limitada; deve ser 1 Gbps+)")
        elif tipo == "Wi-Fi" and not any(other.get("tipo") == "Ethernet" for other in m.get("rede", [])):
            g.append("Operando exclusivamente via Wi-Fi (conexão cabeada Ethernet recomendada para tráfego pesado de arquivos)")

    # Sistema e Energia
    if m.get("plano_energia") == "Economia de Energia":
        g.append("Plano de energia do Windows configurado em 'Economia de Energia' (reduz o clock da CPU e desempenho)")

    uptime = m.get("uptime_dias") or 0
    if uptime > 20:
        g.append(f"Sistema ligado há {uptime} dias sem reiniciar (recomendado reiniciar periodicamente para liberar memória e aplicar atualizações)")

    return g


def nome(m: dict) -> str:
    return (m.get("cadastro") or {}).get("apelido") or m.get("hostname", "?")


def carregar_pasta(pasta: Path) -> list[dict]:
    ms = []
    for f in sorted(pasta.glob("*.json")):
        try:
            d = json.loads(f.read_text(encoding="utf-8"))
            if "hostname" in d and "cpu" in d:
                ms.append(d)
        except Exception:
            pass
    return ms


# ============================================================== RELATÓRIOS
def gerar_docx(maquinas: list[dict], destino: Path) -> None:
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.shared import Pt, Inches

    doc = Document()
    doc.styles["Normal"].font.name = "Segoe UI"
    doc.styles["Normal"].font.size = Pt(9.5)
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = sec.page_height, sec.page_width
    sec.left_margin = Inches(0.6)
    sec.right_margin = Inches(0.6)

    doc.add_heading("Inventário de Hardware e Diagnóstico Técnico", 0)
    doc.add_paragraph(f"Topocart Consultoria · Gerado em {datetime.date.today():%d/%m/%Y} · {len(maquinas)} máquinas catalogadas")

    # 1. Resumo Geral
    doc.add_heading("1. Resumo Geral do Parque de Máquinas", 1)
    cabecalhos = ["Máquina / Hostname", "Responsável / Função", "Processador & Placa-Mãe", "RAM", "GPU & Telas", "Armazenamento & Partições", "Rede & Link"]
    t = doc.add_table(rows=1, cols=len(cabecalhos))
    t.style = "Light Grid Accent 1"
    for c, h in zip(t.rows[0].cells, cabecalhos):
        c.text = h

    for m in maquinas:
        cad = m.get("cadastro") or {}
        linha = t.add_row().cells

        # RAM info
        ram_txt = f"{m.get('ram_gb')} GB {m.get('ram_tipo', '')}\n{m.get('ram_mhz', '')} MHz\n({m.get('ram_canal', '')})"

        # CPU info
        virt = " | VT-x ON" if m.get("virtualizacao_habilitada") else " | VT-x OFF"
        cpu_txt = f"{m['cpu']}\n{m.get('nucleos')}c / {m.get('threads')}t{virt}\nMB: {m.get('placa_mae') or '—'}"

        # GPU & Telas
        telas_txt = ("\nTelas: " + ", ".join(m.get("telas") or [])) if m.get("telas") else ""
        gpu_full = f"{gpu_texto(m)}{telas_txt}"

        # Discos & Volumes
        vols = [v.get("texto", "") for v in m.get("volumes", []) if isinstance(v, dict)]
        discos_txt = "\n".join(vols) if vols else "\n".join(m.get("discos") or [])

        # Rede
        redes_txt = "\n".join(r.get("texto", "") for r in m.get("rede", []) if isinstance(r, dict)) or "—"

        vals = [
            f"{nome(m)}\n({m['hostname']})",
            f"{cad.get('usuario') or m.get('usuario_logado', '')}\n[{cad.get('funcao', 'Geral')}]",
            cpu_txt,
            ram_txt,
            gpu_full,
            discos_txt,
            redes_txt,
        ]
        for c, v in zip(linha, vals):
            c.text = str(v)

    # 2. Pontos de Atenção e Diagnóstico
    doc.add_heading("2. Pontos de Atenção e Diagnóstico de Gargalos", 1)
    criticas = []
    for m in maquinas:
        gs = gargalos(m)
        criticas.append((len(gs), m, gs))
        p = doc.add_paragraph(style="List Bullet")
        p.add_run(f"{nome(m)} ({m['hostname']}): ").bold = True
        p.add_run("; ".join(gs) if gs else "Nenhum ponto de atenção crítico detectado.")
    criticas.sort(key=lambda x: -x[0])

    # 3. Recomendações
    doc.add_heading("3. Recomendações de Otimização e Atualização", 1)
    if criticas and criticas[0][0] > 0:
        piores = ", ".join(f"{nome(m)} ({n} pontos)" for n, m, _ in criticas[:3] if n > 0)
        doc.add_paragraph(
            f"Máquinas prioritárias para upgrade ou remanejamento: {piores}.\n"
            "Principais ações recomendadas:\n"
            "• Ativar Dual Channel adicionando módulos de RAM idênticos nas máquinas em Single Channel.\n"
            "• Substituir ou adicionar unidades SSD NVMe para processamento de nuvens de pontos e projetos pesados.\n"
            "• Verificar cabos/portas de rede limitadas a 100 Mbps para evitar gargalos na transferência de dados.\n"
            "• Ajustar planos de energia para 'Alto Desempenho' e habilitar virtualização (VT-x) nas BIOS quando necessário."
        )
    else:
        doc.add_paragraph("Todas as máquinas analisadas atendem aos parâmetros ideais para fluxos de processamento.")

    doc.add_page_break()

    # 4. Fichas Individuais Detalhadas
    doc.add_heading("4. Fichas Individuais Detalhadas", 1)
    for m in maquinas:
        cad = m.get("cadastro") or {}
        doc.add_heading(f"{nome(m)} — {m['hostname']}", 2)

        pentes_txt = " | ".join(
            f"Slot {p.get('slot')}: {p.get('capacidade_gb')}GB {p.get('tipo')} @ {p.get('velocidade_mhz')}MHz ({p.get('fabricante')})"
            for p in m.get("ram_modulos", [])
        ) or "Não detalhado"

        discos_det = " | ".join(d.get("texto", "") for d in m.get("discos_detalhe", [])) or " | ".join(m.get("discos") or [])
        vols_det = " | ".join(v.get("texto", "") for v in m.get("volumes", [])) or f"C: {m.get('disco_sistema_livre_gb')} GB livres de {m.get('disco_sistema_total_gb')} GB"
        redes_det = " | ".join(f"{r.get('nome')} ({r.get('tipo')} {r.get('link_speed')} - IP: {r.get('ipv4')} - MAC: {r.get('mac')})" for r in m.get("rede", [])) or "—"

        ficha = [
            ("Responsável / Usuário", cad.get("usuario") or m.get("usuario_logado")),
            ("Função Principal", cad.get("funcao") or "Geral"),
            ("Observações do Cadastro", cad.get("obs") or "—"),
            ("Fabricante / Modelo da Máquina", f"{m.get('fabricante') or ''} {m.get('modelo') or ''}".strip()),
            ("Número de Série", m.get("serial") or "—"),
            ("Placa-Mãe", m.get("placa_mae") or "—"),
            ("BIOS (Versão e Data)", f"{m.get('bios_versao') or '—'} (Lançamento: {m.get('bios_data') or '—'})"),
            ("Sistema Operacional", m.get("so")),
            ("Tempo de Atividade (Uptime)", f"{m.get('uptime_dias', 0)} dias sem reiniciar"),
            ("Plano de Energia", m.get("plano_energia") or "—"),
            ("Processador", f"{m['cpu']} ({m.get('nucleos')} núcleos / {m.get('threads')} threads)"),
            ("Virtualização na BIOS (VT-x/AMD-V)", "Habilitada" if m.get("virtualizacao_habilitada") else "Desabilitada"),
            ("Memória RAM Total", f"{m.get('ram_gb')} GB {m.get('ram_tipo', '')} @ {m.get('ram_mhz', '')} MHz — {m.get('ram_canal', '')}"),
            ("Módulos de RAM por Slot", pentes_txt),
            ("Placas de Vídeo (GPU)", gpu_texto(m)),
            ("Monitores Conectados", " | ".join(m.get("telas") or []) or "Padrão"),
            ("Discos Físicos (Hardware)", discos_det),
            ("Partições e Espaço em Disco", vols_det),
            ("Interfaces de Rede e Conectividade", redes_det),
            ("Pontos de Atenção", "; ".join(gargalos(m)) or "Nenhum"),
            ("Data e Hora da Coleta", m.get("data_coleta")),
        ]

        tt = doc.add_table(rows=0, cols=2)
        tt.style = "Light List Accent 1"
        for k, v in ficha:
            r = tt.add_row().cells
            r[0].text = k
            r[1].text = str(v or "—")

    doc.save(destino)


def gerar_xlsx(maquinas: list[dict], destino: Path) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = Workbook()
    ws = wb.active
    ws.title = "Inventário Completo"

    cab = [
        "Apelido", "Hostname", "Responsável", "Função", "Fabricante", "Modelo", "Placa-Mãe", "BIOS Versão",
        "SO / Build", "Uptime (dias)", "Plano Energia", "CPU", "Núcleos", "Threads", "VT-x",
        "RAM (GB)", "RAM Tipo/MHz", "RAM Canal", "Slots Usados", "GPU", "GPU Dedicada",
        "Telas", "Discos Físicos", "Partições / Livre", "Rede (Tipo / Speed / IP)", "Pontos de Atenção", "Data Coleta"
    ]
    ws.append(cab)

    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF", name="Segoe UI", size=10)
        c.fill = PatternFill("solid", fgColor="1E4A8A")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for m in maquinas:
        cad = m.get("cadastro") or {}
        vols_txt = " | ".join(v.get("texto", "") for v in m.get("volumes", [])) or f"C: {m.get('disco_sistema_livre_gb')}GB"
        redes_txt = " | ".join(f"{r.get('tipo')} {r.get('link_speed')} ({r.get('ipv4')})" for r in m.get("rede", [])) or "—"
        ram_speed_str = f"{m.get('ram_tipo', '')} {m.get('ram_mhz', '')}MHz".strip()

        ws.append([
            nome(m),
            m["hostname"],
            cad.get("usuario") or m.get("usuario_logado"),
            cad.get("funcao"),
            m.get("fabricante"),
            m.get("modelo"),
            m.get("placa_mae"),
            m.get("bios_versao"),
            m.get("so"),
            m.get("uptime_dias"),
            m.get("plano_energia"),
            m["cpu"],
            m.get("nucleos"),
            m.get("threads"),
            "Sim" if m.get("virtualizacao_habilitada") else "Não",
            m.get("ram_gb"),
            ram_speed_str,
            m.get("ram_canal"),
            m.get("ram_slots_usados"),
            gpu_texto(m),
            "Sim" if tem_gpu_dedicada(m) else "Não",
            " | ".join(m.get("telas") or []),
            " | ".join(m.get("discos") or []),
            vols_txt,
            redes_txt,
            "; ".join(gargalos(m)),
            m.get("data_coleta"),
        ])

    for col in ws.columns:
        max_len = max(len(str(c.value or "")) for c in col)
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    ws.freeze_panes = "A2"
    wb.save(destino)


def gerar_relatorio(pasta: Path) -> tuple[Path, Path | None, int]:
    maquinas = carregar_pasta(pasta)
    if not maquinas:
        raise RuntimeError("Nenhum JSON de inventário encontrado na pasta.")
    base = pasta / f"Relatorio_Hardware_{datetime.date.today():%Y%m%d}"
    docx = base.with_suffix(".docx")
    gerar_docx(maquinas, docx)
    xlsx = None
    try:
        xlsx = base.with_suffix(".xlsx")
        gerar_xlsx(maquinas, xlsx)
    except ImportError:
        xlsx = None
    return docx, xlsx, len(maquinas)


# ============================================================== INTERFACE GRÁFICA
def main() -> None:
    from PySide6.QtCore import QObject, QThread, Signal
    from PySide6.QtWidgets import (
        QApplication, QFileDialog, QFormLayout, QHBoxLayout, QLabel,
        QLineEdit, QMessageBox, QPlainTextEdit, QPushButton, QVBoxLayout, QWidget
    )

    STYLE = """
    QWidget { font-family: 'Segoe UI', sans-serif; font-size: 10pt; }
    QLineEdit, QPlainTextEdit { border: 1px solid #DDE1E7; border-radius: 4px; padding: 6px; background: white; }
    QPushButton { padding: 8px 16px; border: 1px solid #DDE1E7; border-radius: 4px; background: white; font-weight: 500; }
    QPushButton:hover { background: #ECEEF1; }
    QPushButton#primary { background: #1E4A8A; color: white; font-weight: bold; border: none; }
    QPushButton#primary:hover { background: #153564; }
    QPushButton:disabled { color: #8B95A1; background: #F5F6F7; }
    QLabel#hint { color: #64748B; font-size: 9pt; }
    """

    class Worker(QObject):
        pronto = Signal(object)
        erro = Signal(str)

        def __init__(self, cad: dict):
            super().__init__()
            self.cad = cad

        def run(self):
            try:
                self.pronto.emit(coletar(self.cad))
            except Exception as e:
                self.erro.emit(str(e))

    class Janela(QWidget):
        def __init__(self):
            super().__init__()
            self.setWindowTitle("Inventário de Hardware & Diagnóstico de Sistema — Topocart")
            self.resize(720, 680)
            lay = QVBoxLayout(self)

            lay.addWidget(QLabel("<b>1. Identificação desta Máquina</b>"))
            f = QFormLayout()
            self.e_apelido = QLineEdit()
            self.e_user = QLineEdit(getpass.getuser())
            self.e_func = QLineEdit()
            self.e_obs = QLineEdit()
            self.e_apelido.setPlaceholderText("Ex: Workstation-01, Notebook-Matheus")
            self.e_func.setPlaceholderText("Ex: Processamento LiDAR, CAD, Fotogrametria, Administrativo")

            f.addRow("Apelido da Máquina:", self.e_apelido)
            f.addRow("Responsável:", self.e_user)
            f.addRow("Função Principal:", self.e_func)
            f.addRow("Observações:", self.e_obs)
            lay.addLayout(f)

            lay.addWidget(QLabel("<b>2. Pasta de Inventário</b> (pasta compartilhada na rede ou local)"))
            r = QHBoxLayout()
            self.e_pasta = QLineEdit(str(Path.cwd() / "inventario"))
            b = QPushButton("Selecionar...")
            b.clicked.connect(self.escolher)
            r.addWidget(self.e_pasta)
            r.addWidget(b)
            lay.addLayout(r)

            r_btn = QHBoxLayout()
            self.b_an = QPushButton("▶  Analisar e Salvar Esta Máquina")
            self.b_an.setObjectName("primary")
            self.b_an.clicked.connect(self.analisar)
            self.b_rel = QPushButton("Gerar Relatório Consolidado (.docx / .xlsx)")
            self.b_rel.clicked.connect(self.relatorio)
            r_btn.addWidget(self.b_an)
            r_btn.addWidget(self.b_rel)
            lay.addLayout(r_btn)

            self.out = QPlainTextEdit()
            self.out.setReadOnly(True)
            lay.addWidget(self.out, 1)

            h = QLabel(
                "Instrução: Execute 'Analisar e Salvar' em cada computador apontando para a mesma pasta de rede compartilhada.\n"
                "Em seguida, clique em 'Gerar Relatório Consolidado' em qualquer máquina para compilar os dados em .docx e .xlsx."
            )
            h.setObjectName("hint")
            h.setWordWrap(True)
            lay.addWidget(h)

        def escolher(self):
            d = QFileDialog.getExistingDirectory(self, "Selecionar Pasta de Inventário", self.e_pasta.text())
            if d:
                self.e_pasta.setText(d)

        def analisar(self):
            cad = {
                "apelido": self.e_apelido.text().strip(),
                "usuario": self.e_user.text().strip(),
                "funcao": self.e_func.text().strip(),
                "obs": self.e_obs.text().strip(),
            }
            self.b_an.setEnabled(False)
            self.b_an.setText("Coletando diagnóstico completo…")
            self.out.setPlainText("Coletando informações de hardware, sistema, memória, armazenamento e rede...")

            self.th = QThread(self)
            self.w = Worker(cad)
            self.w.moveToThread(self.th)
            self.th.started.connect(self.w.run)
            self.w.pronto.connect(self.concluido)
            self.w.erro.connect(lambda m: (self.out.appendPlainText(f"ERRO: {m}"), self.fim()))
            self.th.start()

        def fim(self):
            self.th.quit()
            self.th.wait()
            self.b_an.setEnabled(True)
            self.b_an.setText("▶  Analisar e Salvar Esta Máquina")

        def concluido(self, m: dict):
            self.fim()
            pasta = Path(self.e_pasta.text())
            try:
                pasta.mkdir(parents=True, exist_ok=True)
                arquivo_json = pasta / f"{m['hostname']}.json"
                arquivo_json.write_text(json.dumps(m, ensure_ascii=False, indent=2), encoding="utf-8")
                salvo = f"Salvo com sucesso em: {arquivo_json}"
            except Exception as e:
                salvo = f"[ERRO AO SALVAR ARQUIVO]: {e}"

            gs = gargalos(m)
            vols_resumo = " | ".join(v.get("texto", "") for v in m.get("volumes", [])) or f"C: {m.get('disco_sistema_livre_gb')} GB livres"
            redes_resumo = " | ".join(r.get("texto", "") for r in m.get("rede", [])) or "—"
            virt_txt = "Habilitada" if m.get("virtualizacao_habilitada") else "Desabilitada"

            linhas = [
                f"=== DIAGNÓSTICO CONCLUÍDO: {nome(m)} ({m['hostname']}) ===",
                f"Sistema Operacional: {m.get('so')}",
                f"Plano de Energia: {m.get('plano_energia')} | Uptime: {m.get('uptime_dias', 0)} dias",
                f"Placa-Mãe: {m.get('placa_mae') or '—'} | BIOS: {m.get('bios_versao')} ({m.get('bios_data')})",
                f"Processador: {m['cpu']} ({m.get('nucleos')}c / {m.get('threads')}t) | Virtualização VT-x: {virt_txt}",
                f"Memória RAM: {m.get('ram_gb')} GB {m.get('ram_tipo')} @ {m.get('ram_mhz')} MHz ({m.get('ram_canal')}) - {m.get('ram_slots_usados')} slot(s) ocupado(s)",
                f"Placa de Vídeo (GPU): {gpu_texto(m)}",
                f"Monitores/Telas: {' | '.join(m.get('telas') or []) or 'Padrão'}",
                f"Discos Físicos: {' | '.join(m.get('discos') or [])}",
                f"Partições / Espaço: {vols_resumo}",
                f"Conectividade de Rede: {redes_resumo}",
                "",
                "PONTOS DE ATENÇÃO DETECTADOS:",
                "  • " + ("\n  • ".join(gs) if gs else "Nenhum gargalo identificado (máquina bem equilibrada)."),
                "",
                salvo,
            ]
            self.out.setPlainText("\n".join(linhas))

        def relatorio(self):
            try:
                docx, xlsx, n = gerar_relatorio(Path(self.e_pasta.text()))
            except Exception as e:
                QMessageBox.warning(self, "Aviso de Relatório", str(e))
                return
            self.out.setPlainText(
                f"Relatório gerado com sucesso para {n} máquina(s)!\n\n"
                f"Documento Word (.docx): {docx}\n"
                f"Planilha Excel (.xlsx): {xlsx or '(openpyxl não instalado)'}"
            )

    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(STYLE)
    j = Janela()
    j.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
