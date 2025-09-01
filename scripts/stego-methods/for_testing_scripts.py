from pathlib import Path

base = "data_set/stego-files"
for directories in Path(base).iterdir():
    print(directories.name)
    for file in directories.iterdir():
        if file.is_file():
            # TE: apstrāde konkrētajai apakšmapei
            docPath = f"{base}/{directories.name}/{file.name}"
            print("  fails:", docPath)
