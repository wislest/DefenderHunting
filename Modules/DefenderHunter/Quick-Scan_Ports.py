import socket
from concurrent.futures import ThreadPoolExecutor
import time

def test_port(ip, port, timeout=1):
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(timeout)
    try:
        result = sock.connect_ex((ip, port))
        if result == 0:
            return port, "ouvert"
        return port, "fermé"
    except:
        return port, "erreur"
    finally:
        sock.close()

# Ports courants et fréquemment ciblés
target_ports = {
    20: "FTP Data",
    21: "FTP Control",
    22: "SSH",
    23: "Telnet",
    25: "SMTP",
    53: "DNS",
    69: "TFTP",
    80: "HTTP",
    88: "Kerberos",
    110: "POP3",
    111: "RPCBind",
    123: "NTP",
    135: "MSRPC",
    137: "NetBIOS",
    139: "NetBIOS Session",
    143: "IMAP",
    161: "SNMP",
    389: "LDAP",
    443: "HTTPS",
    445: "SMB",
    623: "IPMI",
    631: "IPP",
    1433: "MSSQL",
    1434: "MSSQL Browser",
    1521: "Oracle",
    2049: "NFS",
    3306: "MySQL",
    3389: "RDP",
    4444: "Common Malware",
    5432: "PostgreSQL",
    5900: "VNC",
    6379: "Redis",
    8080: "HTTP Alt",
    8443: "HTTPS Alt",
    9200: "Elasticsearch",
    27017: "MongoDB"
}

ip = "72.138.161.61"

print(f"Scan étendu des ports pour {ip}...")
print("(inclut les ports couramment vulnérables et attaqués)")
start_time = time.time()

with ThreadPoolExecutor(max_workers=15) as executor:
    futures = [executor.submit(test_port, ip, port) for port in target_ports.keys()]
    results = []
    for future in futures:
        port, status = future.result()
        results.append((port, status))
    
    # Trier les résultats par numéro de port
    results.sort(key=lambda x: x[0])
    
    # Afficher les résultats
    for port, status in results:
        if status == "ouvert":
            print(f"Port {port:5d} ({target_ports[port]:15s}): {status} ⚠️")
        else:
            print(f"Port {port:5d} ({target_ports[port]:15s}): {status}")

print(f"\nScan terminé en {time.time() - start_time:.2f} secondes")