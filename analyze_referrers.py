import csv
from collections import Counter
import os

def analyze_referrers():
    input_path = './output/future-appointments_REAL.csv'
    output_path = './output/referrer_analysis.csv'
    
    if not os.path.exists(input_path):
        print(f"Error: {input_path} not found.")
        return

    referrers = []
    try:
        with open(input_path, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                # Some files might have leading/trailing whitespace in headers
                # Try to find a header that matches 'referring_doctor'
                ref_key = None
                for k in row.keys():
                    if k.strip().lower() == 'referring_doctor':
                        ref_key = k
                        break
                
                if ref_key:
                    ref = row[ref_key].strip()
                    if ref:
                        # Clean name: remove "Dr" prefix and title case
                        if ref.lower().startswith('dr '):
                            ref = ref[3:].strip()
                        elif ref.lower().startswith('dr.'):
                            ref = ref[3:].strip()
                        
                        ref = ref.title()
                        referrers.append(ref)
                else:
                    # Fallback or debug
                    pass

        # Count occurrences
        counts = Counter(referrers).most_common()
        
        # Save to CSV
        with open(output_path, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Referrer', 'Number of Referrals'])
            for name, count in counts:
                writer.writerow([f"Dr {name}", count])
        
        print(f"Analysis complete. Results saved to {output_path}")
        print("\n--- Top Referrers ---")
        for name, count in counts[:10]:
            print(f"Dr {name}: {count}")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    analyze_referrers()
