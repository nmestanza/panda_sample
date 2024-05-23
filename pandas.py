import paramiko
import getpass
import openpyxl

def get_ap_count(hostname, username, password):
    try:
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(hostname, username=username, password=password)

        stdin, stdout, stderr = ssh.exec_command('show ap summary')
        output = stdout.read().decode()
        
        ssh.close()
        
        # Process the output to find the number of APs
        lines = output.splitlines()
        ap_count = 0
        for line in lines:
            if 'Number of APs' in line:
                ap_count = int(line.split()[-1])
                break
        
        return ap_count

    except Exception as e:
        print(f"Failed to connect to {hostname}: {e}")
        return None

def main():
    # Ask for credentials
    username = input("Enter your SSH username: ")
    password = getpass.getpass("Enter your SSH password: ")
    
    # Read hostnames from the input file
    with open('hostnames.txt', 'r') as file:
        hostnames = [line.strip() for line in file.readlines()]
    
    # Create a new Excel workbook and select the active worksheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'AP Counts'
    
    # Write headers
    sheet['A1'] = 'Hostname'
    sheet['B1'] = 'AP Count'
    
    # Collect AP counts for each hostname
    for row, hostname in enumerate(hostnames, start=2):
        print(f"Connecting to {hostname}...")
        ap_count = get_ap_count(hostname, username, password)
        if ap_count is not None:
            sheet[f'A{row}'] = hostname
            sheet[f'B{row}'] = ap_count
        else:
            sheet[f'A{row}'] = hostname
            sheet[f'B{row}'] = 'Error'
    
    # Save the workbook to a file
    workbook.save('ap_counts.xlsx')
    print("AP counts have been saved to 'ap_counts.xlsx'")

if __name__ == "__main__":
    main()