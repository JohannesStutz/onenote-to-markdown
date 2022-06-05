from convert import safe_str

assert safe_str("test") == "test"
assert safe_str("hä hä") == "hä hä"
assert safe_str(r'file<>:"/\|?*name') == "file" + "_" * 9 + "name"
assert safe_str("CON") == "CON_note"
assert safe_str("LPT1") == "LPT1_note"
