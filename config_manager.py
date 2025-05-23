import json
import os
from utils import log

class ConfigManager:
    """Handles loading, saving, and managing application configuration."""
    def __init__(self, app_dir, default_config):
        self.config_file = os.path.join(app_dir, "config.json")
        self.default_config = default_config
        self.config = self._load_config()

    def _load_config(self):
        """Loads application configuration from file, merging with defaults."""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, "r", encoding="utf-8") as f:
                    loaded_config = json.load(f)
                    config = self.default_config.copy()
                    config.update(loaded_config)
                    return config
        except Exception as e:
            log(f"Error loading config: {e}")
            if os.path.exists(self.config_file):
                backup_file = self.config_file + ".bak"
                try:
                    import shutil
                    shutil.copy2(self.config_file, backup_file)
                    log(f"Created backup of corrupted config: {backup_file}")
                except Exception as e_bak:
                    log(f"Error creating config backup: {e_bak}")
        return self.default_config.copy()

    def save_config(self):
        """Saves current application configuration to file."""
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(self.config, f, ensure_ascii=False, indent=4)
            return True
        except Exception as e:
            log(f"Error saving config: {e}")
            return False

    def get(self, key, default=None):
        """Retrieves a configuration value."""
        return self.config.get(key, default)

    def set(self, key, value):
        """Sets a configuration value and saves the config."""
        self.config[key] = value
        self.save_config()

    def reset_to_default(self):
        """Resets all configuration settings to their default values."""
        self.config = self.default_config.copy()
        return self.save_config()