import json
import os
from utils import log

class SnippetManager:
    """Handles loading, saving, and managing text snippets."""
    def __init__(self, app_dir):
        self.snippets_file = os.path.join(app_dir, "snippets.json")
        self.snippets = self._load_snippets()

    def _load_snippets(self):
        """Loads snippets from file, including format conversion and error handling."""
        try:
            if os.path.exists(self.snippets_file):
                with open(self.snippets_file, "r", encoding="utf-8") as f:
                    snippets = json.load(f)
                    updated = False
                    for shortcut, data in snippets.items():
                        if isinstance(data, str):
                            snippets[shortcut] = {
                                "text": data,
                                "category": "General",
                                "description": "",
                            }
                            updated = True
                    if updated:
                        with open(self.snippets_file, "w", encoding="utf-8") as f:
                            json.dump(snippets, f, ensure_ascii=False, indent=4)
                    return snippets
        except Exception as e:
            log(f"Error loading snippets: {e}")
            if os.path.exists(self.snippets_file):
                backup_file = self.snippets_file + ".bak"
                try:
                    import shutil
                    shutil.copy2(self.snippets_file, backup_file)
                    log(f"Created backup of corrupted snippets: {backup_file}")
                except Exception as e_bak:
                    log(f"Error creating snippet backup: {e_bak}")
        return self._get_default_snippets()

    def _get_default_snippets(self):
        """Returns a dictionary of default snippets."""
        return {
            "/date": {
                "text": "{date}",
                "category": "Date & Time",
                "description": "Insert current date",
            },
            "/time": {
                "text": "{time}",
                "category": "Date & Time",
                "description": "Insert current time",
            },
            "/datetime": {
                "text": "{date} {time}",
                "category": "Date & Time",
                "description": "Insert date and time",
            },
            "/sig": {
                "text": "Best regards,\n{input:Your Name?}",
                "category": "Email",
                "description": "Email signature",
            },
            "/addr": {
                "text": "123 Main St\nAnytown, CA 90210",
                "category": "Personal",
                "description": "Sample mailing address",
            },
            "/meeting": {
                "text": "Meeting: {input:Meeting title?}\nDate: {date}\nTime: {time}\nAttendees: {input:Attendees?}",
                "category": "Work",
                "description": "Meeting template with dynamic inputs",
            },
        }

    def save_snippets(self):
        """Saves current snippets to file."""
        try:
            with open(self.snippets_file, "w", encoding="utf-8") as f:
                json.dump(self.snippets, f, ensure_ascii=False, indent=4)
            return True
        except Exception as e:
            log(f"Error saving snippets: {e}")
            return False

    def get_all_snippets(self):
        """Returns all loaded snippets."""
        return self.snippets

    def get_snippet(self, shortcut):
        """Returns snippet data for a given shortcut."""
        return self.snippets.get(shortcut)

    def add_update_snippet(self, shortcut, text, category="General", description=""):
        """Adds or updates a snippet."""
        self.snippets[shortcut] = {
            "text": text,
            "category": category,
            "description": description,
        }
        return self.save_snippets()

    def delete_snippet(self, shortcut):
        """Deletes a snippet by shortcut."""
        if shortcut in self.snippets:
            del self.snippets[shortcut]
            return self.save_snippets()
        return False

    def get_all_categories(self):
        """Returns a sorted list of all unique categories from snippets."""
        categories = set()
        for data in self.snippets.values():
            if isinstance(data, dict) and "category" in data:
                categories.add(data["category"])
            else:
                categories.add("General")
        categories.add("General")
        return sorted(list(categories))