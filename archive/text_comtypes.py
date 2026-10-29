from comtypes.client import CreateObject

try:
    print("Testing comtypes with ProgID...")
    swApp = CreateObject("DSComposer.Application")
    print("Connected to SOLIDWORKS Composer!")
    # Try type library methods
    guesses = ["GetActiveDocument", "ActiveDoc", "GetDocument", "Open"]
    for guess in guesses:
        try:
            if guess == "Open":
                swDoc = getattr(swApp, guess)("C:\OneDrive - Epiroc\Product Information Library (by Skyfacture)\7394 1000 38 RD314 38mm DRILLING\7394 1000 38.smg")  # Replace with real path
            elif guess == "GetDocument":
                swDoc = getattr(swApp, guess)()  # Guess doc ID
            else:
                swDoc = getattr(swApp, guess)()
            if swDoc:
                print(f"{guess} worked! Views: {swDoc.ViewCount}")
                break
        except AttributeError:
            print(f"{guess} not found.")
    # Dump methods if all fail
    print("Available methods:", dir(swApp))
except Exception as e:
    print(f"Error: {repr(e)}")