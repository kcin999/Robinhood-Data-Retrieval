import toml

cfgG = "config.toml"

def get_robinhood_username():
	cfg = toml.load(cfgG)
	value = cfg.get("robinhood",{}).get("username")
	if value is None:
		raise Exception ("Can not parse library name from config.toml.\nconfig.toml missing value for 'username' key\nin the [robinhood] section")
	return value

def get_robinhood_password():
	cfg = toml.load(cfgG)
	value = cfg.get("robinhood",{}).get("password")
	if value is None:
		raise Exception ("Can not parse library name from config.toml.\nconfig.toml missing value for 'password' key\nin the [robinhood] section")
	return value

def get_robinhood_url():
	cfg = toml.load(cfgG)
	value = cfg.get("robinhood",{}).get("url")
	if value is None:
		raise Exception ("Can not parse library name from config.toml.\nconfig.toml missing value for 'url' key\nin the [robinhood] section")
	return value

def get_robinhood_mfa_code():
	cfg = toml.load(cfgG)
	value = cfg.get("robinhood",{}).get("mfa_code")
	if value is None:
		raise Exception ("Can not parse library name from config.toml.\nconfig.toml missing value for 'mfa_code' key\nin the [robinhood] section")
	return value

def get_excel_fileName():
	cfg = toml.load(cfgG)
	value = cfg.get("excel",{}).get("fileName")
	if value is None:
		raise Exception ("Can not parse library name from config.toml.\nconfig.toml missing value for 'fileName' key\nin the [excel] section")
	return value

def get_csv_fileName():
	cfg = toml.load(cfgG)
	value = cfg.get("csv",{}).get("fileName")
	if value is None:
		raise Exception ("Can not parse library name from config.toml.\nconfig.toml missing value for 'fileName' key\nin the [csv] section")
	return value