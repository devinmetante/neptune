from inspect import getmembers

import win32com.client

progID = 'Neptune.NeptuneCtrl.1'
# progID = '{50182A02-9665-4094-AA35-B1938D210DAC}'
w = win32com.client.gencache.EnsureDispatch(progID)


def print_members(obj, obj_name):
    """Print members of given COM object"""
    try:
        fields = list(obj._prop_map_get_.keys())
    except AttributeError:
        print("Object has no attribute '_prop_map_get_'")
        print("Check if the initial COM object was created with"
              "'win32com.client.gencache.EnsureDispatch()'")
        raise
    methods = [m[0] for m in getmembers(obj) if (not m[0].startswith("_")
                                                 and "clsid" not in m[0].lower())]

    if len(fields) + len(methods) > 0:
        print("Members of '{}' ({}):".format(obj_name, obj))
    else:
        raise ValueError("Object has no members to print")

    print("\tFields:")
    if fields:
        for field in fields:
            print(f"\t\t{field}")
    else:
        print("\t\tObject has no fields to print")

    print("\tMethods:")
    if methods:
        for method in methods:
            print(f"\t\t{method}")
    else:
        print("\t\tObject has no methods to print")


print_members(w, progID)
