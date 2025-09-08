import os

source = r"G:\Shared drives\Data Management\1_Daily Operation\3. 2D & 3D files\19_VCT & VCC"
target = r"D:\VCT\2D&3D"

def get_subfolders(path):
    if not os.path.isdir(path):
        return set()
    return {f for f in os.listdir(path) if os.path.isdir(os.path.join(path, f))}

# หาโฟลเดอร์หลักที่มีทั้งสองฝั่ง (intersect)
source_folders = {f for f in os.listdir(source) if os.path.isdir(os.path.join(source, f))}
target_folders = {f for f in os.listdir(target) if os.path.isdir(os.path.join(target, f))}
common_folders = source_folders & target_folders

print(f"พบโฟลเดอร์ที่ทั้งสองฝั่งมีเหมือนกัน {len(common_folders)} โฟลเดอร์")

for folder in sorted(common_folders):
    src_path = os.path.join(source, folder)
    tgt_path = os.path.join(target, folder)

    src_sub = get_subfolders(src_path)
    tgt_sub = get_subfolders(tgt_path)

    if src_sub != tgt_sub:
        print(f"\n❌ โฟลเดอร์ '{folder}' มี subfolder ไม่ตรงกัน:")
        print(f"   Source ({source}): {src_sub}")
        print(f"   Target ({target}): {tgt_sub}")
