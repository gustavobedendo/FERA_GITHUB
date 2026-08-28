import configparser
import os
CONFIG_FILE = 'settings.ini'
SECTION = 'Preferences'
OPTION  = 'dontshowagain'
def on_toggle(var):
    config.set(SECTION, OPTION, str(var.get()))
    with open(CONFIG_FILE, 'w') as f:
        config.write(f)
def get_dontshow():
    return config.getint(SECTION, OPTION, fallback=0)
# 1) Load existing setting (or default to 0)
config = configparser.ConfigParser()
config.read(CONFIG_FILE)
if not config.has_section(SECTION):
    config.add_section(SECTION)
dont_show = config.getint(SECTION, OPTION, fallback=0)