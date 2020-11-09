import json

class ConfigsManager():
    file_name = "strava_logger.cfg"

    def read_configs(self):
        configs = {}
        with open(self.file_name) as json_file:
            configs = json.load(json_file)
            print("configs= " + str(configs))
        return configs

    def write_configs(self, configs):
        with open(self.file_name, 'w') as outfile:
            json.dump(configs, outfile)