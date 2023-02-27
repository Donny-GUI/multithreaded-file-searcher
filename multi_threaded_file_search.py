import os
import threading
from queue import Queue
import json
import csv 


USER_HOME = os.path.expanduser('~')


class FileFinder:
    def __init__(self, start_dir: str, file_extension: str) -> None:
        self.start_dir = start_dir
        self.file_extension = file_extension
        self.result_queue = Queue()
        self.threads = []
        self.subdirs = [os.path.join(self.start_dir, subdir) for subdir in os.listdir(self.start_dir) if os.path.isdir(os.path.join(self.start_dir, subdir))]
        self.subdirs_count = len(self.subdirs)
        
    def find_files(self) -> None:
        
        for root, dirs, files in os.walk(self.start_dir):
            for file in files:
                if file.endswith(".xlsx") or file.endswith(".xls"):
                    self.result_queue.put(os.path.join(root, file))

    def start(self, num_threads: int) -> None:
        
        for subdir in self.subdirs:
            thread = threading.Thread(target=self.find_excel_files, args=(subdir,))
            thread.start()
            self.threads.append(thread)
        for thread in self.threads:
            thread.join()

    def get_results(self) -> None:
        
        results = []
        while not self.result_queue.empty():
            results.append(self.result_queue.get())
        return results

class FileSearch(FileFinder):
    def __init__(self, start_dir: str=USER_HOME, file_extension: str= '.xlsx', output_file: str="", output_type: str="") -> None:
        super().__init__(start_dir, file_extension)
        
        if not output_type.startswith('.') and output_type != "":
            output_type = "." + output_type
        if not output_file == "" and not output_file.endswith(output_type):
            output_file = output_file + output_type
        self.output_file = output_file
        self.output_type = output_type
        self.output_plz = True if self.output_file != "" else False
        self.start(self.subdirs_count)
        self.output() if self.output_plz else self.passer()  
    
    def passer(self):
        
        for file in self.get_results():
            print(file)
            
    def output(self):
        
        if self.output_type == '.json':
            with open(self.output_file, 'w') as jfile:
                json.dump(self.get_results())
            self.passer()
        elif self.output_type == '.csv':
            with open(self.output_file, "w", newline="") as f:
                writer = csv.writer(f)
                for row in self.get_results():
                    writer.writerow(row)
            self.passer()
            
        