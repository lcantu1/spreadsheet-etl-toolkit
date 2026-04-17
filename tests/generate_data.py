import csv
import random
import argparse
from faker import Faker

def generate_analytics_data(filename, num_rows):
    """Generates mock web analytics data to test ETL and deduplication scripts."""
    fake = Faker()
    
    # A constrained list of base paths to intentionally create duplicate rows for testing
    base_paths = [
        ("/shop/electronics/monitors", "Shop", "Electronics", "Product"),
        ("/shop/home-goods/chairs", "Shop", "Furniture", "Product"),
        ("/blog/tech/homelab-setup", "Blog", "Tech", "Article"),
        ("/blog/recipes/savory-sweet-beef", "Blog", "Food", "Article"),
        ("/support/faq/shipping", "Support", "FAQ", "Help"),
        ("/about-us", "Corporate", "About", "Info")
    ]
    
    print(f"🚀 Generating {num_rows} rows of mock data...")
    
    with open(filename, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        
        # Write the exact headers your Google Apps Script expects
        writer.writerow(['Data_ID', 'Page path', 'Total users', 'Primary_Category', 'Sub_Category', 'Page_Type_Normalized', 'Target_URL'])
        
        for _ in range(num_rows):
            # Generate a completely unique Data ID for every row
            data_id = f"DAT-{fake.unique.random_int(min=10000, max=999999)}"
            
            # Pick a random path from our pool
            path_base, primary, sub, p_type = random.choice(base_paths)
            
            # Simulate real-world messy data
            # 15% chance to inject the buggy '.html.html' extension
            if random.random() < 0.15:
                path = path_base + ".html.html"
            # 40% chance for a standard .html, 45% chance for a trailing slash
            elif random.random() < 0.40:
                path = path_base + ".html"
            else:
                path = path_base + "/"
                
            # Construct the full URL
            url = f"https://example.com{path}"
            
            # Generate random user traffic volume
            users = random.randint(10, 5500)
                
            writer.writerow([data_id, path, users, primary, sub, p_type, url])
            
    print(f"✅ Success! Data saved to {filename}")

if __name__ == "__main__":
    # Set up command-line arguments so the script operates like a professional CLI tool
    parser = argparse.ArgumentParser(description="Generate mock web analytics data.")
    parser.add_argument('--rows', type=int, default=100, help='Number of rows to generate (default: 100)')
    parser.add_argument('--output', type=str, default='mock_dataset.csv', help='Output CSV filename')
    
    args = parser.parse_args()
    
    generate_analytics_data(args.output, args.rows)
