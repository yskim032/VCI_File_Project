import pandas as pd
from collections import defaultdict
from typing import List, Dict, Set

class ContainerAnalyzer:
    def __init__(self, operation_type: str, tpf_containers: Set[str], truck_containers: Set[str]):
        self.operation_type = operation_type  # 'DIS' or 'LOD'
        self.tpf_containers = tpf_containers
        self.truck_containers = truck_containers
        self.container_groups = defaultdict(list)

    def parse_container_data(self, line: str) -> Dict:
        """Parse a single line from ASC file"""
        container_number = line[6:18].strip()
        container_type = line[44:48].strip()  # 45-48 position
        full_empty = line[51:52].strip()      # 52 position
        operator_code = line[19:22].strip()   # 20-22 position

        # Create container key based on all fields that affect grouping
        group_key = (
            self.operation_type,
            container_type,
            full_empty,
            operator_code,
            'No',  # OOG
            'No',  # Damaged
            'No',  # SOC
            'No',  # Coastal Cargo
            'No',  # To Rail
            'No',  # To Barge
            'Yes' if container_number in self.tpf_containers else 'No',  # To TPF
            'Yes' if container_number in self.truck_containers else 'No',  # To Truck
            'No'   # Not for MSC Account
        )

        return {
            'container_number': container_number,
            'group_key': group_key,
            'data': {
                'Operation': self.operation_type,
                'Container Type': container_type,
                'Full/Empty': full_empty,
                'Operator Code': operator_code,
                'OOG': 'No',
                'Damaged': 'No',
                'SOC': 'No',
                'Coastal Cargo': 'No',
                'To Rail': 'No',
                'To Barge': 'No',
                'To TPF': 'Yes' if container_number in self.tpf_containers else 'No',
                'To Truck': 'Yes' if container_number in self.truck_containers else 'No',
                'Not for MSC Account': 'No'
            }
        }

    def process_file(self, file_path: str) -> pd.DataFrame:
        """Process ASC file and return summary DataFrame"""
        # Read and process the file
        with open(file_path, 'r', encoding='utf-8') as f:
            for line in f:
                if not line.startswith('$'):  # Skip header lines
                    container_data = self.parse_container_data(line)
                    self.container_groups[container_data['group_key']].append(
                        container_data['container_number']
                    )

        # Create summary records
        summary_records = []
        for group_key, containers in self.container_groups.items():
            record = {
                'Operation': group_key[0],
                'Container Type': group_key[1],
                'Full/Empty': group_key[2],
                'Operator Code': group_key[3],
                'Quantity': len(containers),
                'OOG': group_key[4],
                'Damaged': group_key[5],
                'SOC': group_key[6],
                'Coastal Cargo': group_key[7],
                'To Rail': group_key[8],
                'To Barge': group_key[9],
                'To TPF': group_key[10],
                'To Truck': group_key[11],
                'Not for MSC Account': group_key[12]
            }
            summary_records.append(record)

        return pd.DataFrame(summary_records)

def create_summary(asc_file: str, operation_type: str, 
                  tpf_containers: List[str], truck_containers: List[str], 
                  output_file: str):
    """
    Create container summary Excel file
    
    Args:
        asc_file: Path to ASC file
        operation_type: 'DIS' or 'LOD'
        tpf_containers: List of container numbers for TPF
        truck_containers: List of container numbers for Truck
        output_file: Path to output Excel file
    """
    # Convert container lists to sets for faster lookup
    tpf_set = set(tpf_containers)
    truck_set = set(truck_containers)

    # Create analyzer and process file
    analyzer = ContainerAnalyzer(operation_type, tpf_set, truck_set)
    summary_df = analyzer.process_file(asc_file)

    # Write to Excel
    summary_df.to_excel(output_file, index=False)

# Example usage:
if __name__ == "__main__":
    # These would be provided by the user
    asc_file = "ADFT512EIST_F.ASC"
    operation_type = "DIS"  # or "LOD"
    tpf_containers = []  # List of container numbers for TPF
    truck_containers = []  # List of container numbers for Truck
    output_file = "container_summary.xlsx"

    create_summary(asc_file, operation_type, tpf_containers, truck_containers, output_file) 