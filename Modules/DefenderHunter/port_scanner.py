import socket
from concurrent.futures import ThreadPoolExecutor
import time
import argparse
import csv
import datetime
import json
import os
import sys
import ipaddress

class PortScanner:
    def __init__(self):
        self.known_ports = {
            20: "FTP Data", 21: "FTP Control", 22: "SSH", 23: "Telnet",
            25: "SMTP", 53: "DNS", 69: "TFTP", 80: "HTTP", 88: "Kerberos",
            110: "POP3", 111: "RPCBind", 123: "NTP", 135: "MSRPC",
            137: "NetBIOS", 139: "NetBIOS Session", 143: "IMAP", 161: "SNMP",
            389: "LDAP", 443: "HTTPS", 445: "SMB", 623: "IPMI", 631: "IPP",
            1433: "MSSQL", 1434: "MSSQL Browser", 1521: "Oracle", 2049: "NFS",
            3306: "MySQL", 3389: "RDP", 4444: "Common Malware", 5432: "PostgreSQL",
            5900: "VNC", 6379: "Redis", 8080: "HTTP Alt", 8443: "HTTPS Alt",
            9200: "Elasticsearch", 27017: "MongoDB"
        }
        self.results = []

    def test_port(self, ip, port, timeout=1):
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(timeout)
        try:
            result = sock.connect_ex((ip, port))
            service = self.known_ports.get(port, "Service inconnu")
            if result == 0:
                return {"ip": ip, "port": port, "status": "ouvert", "service": service}
            return {"ip": ip, "port": port, "status": "fermé", "service": service}
        except socket.gaierror:
            return {"ip": ip, "port": port, "status": "erreur DNS", "service": service}
        except socket.error:
            return {"ip": ip, "port": port, "status": "erreur réseau", "service": service}
        finally:
            sock.close()

    def scan_ports(self, ip, ports, timeout=1):
        print(f"\nScan des ports pour {ip}...")
        start_time = time.time()
        scan_results = []

        with ThreadPoolExecutor(max_workers=15) as executor:
            futures = [executor.submit(self.test_port, ip, port, timeout) for port in ports]
            for future in futures:
                result = future.result()
                scan_results.append(result)
                self.results.append(result)
                
                # Affichage en temps réel
                if result["status"] == "ouvert":
                    print(f"Port {result['port']:5d} ({result['service']:15s}): {result['status']} ⚠️")
                else:
                    print(f"Port {result['port']:5d} ({result['service']:15s}): {result['status']}")

        duration = time.time() - start_time
        print(f"\nScan terminé en {duration:.2f} secondes")
        return scan_results

    def export_results(self, format="csv"):
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        if format == "csv":
            filename = f"scan_results_{timestamp}.csv"
            with open(filename, 'w', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=["ip", "port", "status", "service"])
                writer.writeheader()
                writer.writerows(self.results)
        elif format == "json":
            filename = f"scan_results_{timestamp}.json"
            with open(filename, 'w') as f:
                json.dump(self.results, f, indent=4)
        print(f"Résultats exportés dans {filename}")

def parse_port_range(port_string):
    ports = set()
    try:
        parts = port_string.split(',')
        for part in parts:
            if '-' in part:
                start, end = map(int, part.split('-'))
                if start > end:
                    raise ValueError("Plage de ports invalide")
                if start < 1 or end > 65535:
                    raise ValueError("Les ports doivent être entre 1 et 65535")
                ports.update(range(start, end + 1))
            else:
                port = int(part)
                if port < 1 or port > 65535:
                    raise ValueError("Les ports doivent être entre 1 et 65535")
                ports.add(port)
        return sorted(list(ports))
    except ValueError as e:
        raise ValueError(f"Format de port invalide: {str(e)}")

def parse_ip_range(ip_string):
    try:
        if '/' in ip_string:  # CIDR notation
            return list(ipaddress.ip_network(ip_string, strict=False))
        elif '-' in ip_string:  # Range notation
            start_ip, end_ip = ip_string.split('-')
            start = ipaddress.ip_address(start_ip.strip())
            end = ipaddress.ip_address(end_ip.strip())
            return [ipaddress.ip_address(ip) for ip in range(int(start), int(end) + 1)]
        else:  # Single IP
            return [ipaddress.ip_address(ip_string.strip())]
    except ValueError as e:
        raise ValueError(f"Format d'IP invalide: {str(e)}")

def main():
    scanner = PortScanner()
    
    while True:
        print("\n=== Scanner de ports - Menu principal ===")
        print("1. Scanner avec les ports prédéfinis")
        print("2. Scanner des ports spécifiques")
        print("3. Scanner une plage d'IPs")
        print("4. Exporter les résultats")
        print("5. Configuration du scan")
        print("6. Quitter")
        
        choix = input("\nChoisissez une option (1-6): ")
        
        if choix == "6":
            break
            
        if choix in ["1", "2", "3"]:
            if choix in ["1", "2"]:
                ip_input = input("Entrez l'adresse IP à scanner: ")
                try:
                    ips = parse_ip_range(ip_input)
                except ValueError as e:
                    print(f"Erreur: {e}")
                    continue
                
                if choix == "1":
                    ports = scanner.known_ports.keys()
                else:
                    print("\nFormats acceptés:")
                    print("- Port unique: 80")
                    print("- Liste de ports: 80,443,8080")
                    print("- Plage de ports: 80-100")
                    print("- Combinaison: 80,443,1000-1100,8080")
                    ports_input = input("\nPorts à scanner: ")
                    try:
                        ports = parse_port_range(ports_input)
                    except ValueError as e:
                        print(f"Erreur: {e}")
                        continue
                
                for ip in ips:
                    scanner.scan_ports(str(ip), ports)
            
        elif choix == "4":
            if not scanner.results:
                print("Aucun résultat à exporter")
                continue
                
            print("\nFormat d'export:")
            print("1. CSV")
            print("2. JSON")
            format_choix = input("Choisissez le format (1-2): ")
            
            if format_choix == "1":
                scanner.export_results("csv")
            elif format_choix == "2":
                scanner.export_results("json")
            else:
                print("Option invalide")
        
        elif choix == "5":
            print("\nParamètres de configuration à venir...")
            # TODO: Ajouter les options de configuration (timeout, threads, etc.)
            
        else:
            print("Option invalide")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nScan interrompu par l'utilisateur")
    except Exception as e:
        print(f"\nErreur inattendue: {e}")