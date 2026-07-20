import os
import shutil
import site
import sys

def main():
    print("Searching for corrupted scikit-image installation...")
    
    # Collect all potential site-packages paths
    paths = []
    
    # 1. Standard site-packages
    try:
        paths.extend(site.getsitepackages())
    except Exception:
        pass
        
    # 2. User site-packages
    try:
        paths.append(site.getusersitepackages())
    except Exception:
        pass
        
    # 3. sys.path entries containing site-packages
    for p in sys.path:
        if "site-packages" in p and p not in paths:
            paths.append(p)
            
    print(f"Paths to search: {paths}")
    
    deleted_any = False
    for path in paths:
        if not os.path.exists(path):
            continue
            
        for item in os.listdir(path):
            item_path = os.path.join(path, item)
            # Match skimage folder or any scikit_image metadata folders
            if item.lower() == "skimage" or item.lower().startswith("scikit_image"):
                print(f"Found: {item_path}")
                try:
                    if os.path.isdir(item_path):
                        shutil.rmtree(item_path)
                    else:
                        os.remove(item_path)
                    print(f"Successfully deleted: {item_path}")
                    deleted_any = True
                except Exception as e:
                    print(f"Failed to delete {item_path}: {e}")
                    print("Please try running this script as Administrator or manually delete this folder.")

    if deleted_any:
        print("\nSuccessfully cleaned up corrupted scikit-image folders!")
        print("Now you can run: pip install -r requirements.txt")
    else:
        print("\nCould not find any scikit-image folders to delete.")

if __name__ == "__main__":
    main()
