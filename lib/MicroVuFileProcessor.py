import os

import lib
from lib.MicroVuProgram import MicroVuProgram
from lib.Utilities import get_unencoded_file_lines, get_filepath_by_name


class CoonRapidsProcessor:
    bring_part_to_metrology_lines: list[str] = []

    def __init__(self):
        self._load_lines()

    def _load_lines(self):
        bring_to_met_text_filepath = get_filepath_by_name('BringPartToMetrology_text.txt')
        if not bring_to_met_text_filepath:
            raise ProcessorException("Can't find 'BringPartToMetrology_text' file.")

        if lines := get_unencoded_file_lines(bring_to_met_text_filepath):
            self.bring_part_to_metrology_lines = lines
        else:
            raise ProcessorException("Can't find 'BringPartToMetrology_text' file.")

    def _inject_bring_to_metrology_picture(self, micro_vu: MicroVuProgram) -> None:

        if micro_vu.instructions_index == -1:
            return

        bring_to_met_pic_idx = micro_vu.bring_part_to_metrology_index
        micro_vu.insert_line(bring_to_met_pic_idx, self.bring_part_to_metrology_lines[3])
        micro_vu.insert_line(bring_to_met_pic_idx, self.bring_part_to_metrology_lines[2])
        micro_vu.insert_line(bring_to_met_pic_idx, self.bring_part_to_metrology_lines[1])

    def _write_file_to_harddrive(self, micro_vu: MicroVuProgram) -> None:

        try:
            with open(micro_vu.filepath, 'w', encoding='utf-16-le', newline='\r\n') as f:
                for line in micro_vu.file_lines:
                    f.write(f"{line}")
        except:
            pass

    def process_files(self) -> None:
        input_rootpath = lib.Utilities.GetStoredIniValue("Paths", "input_rootpath", "Settings")

        for root, directory, files in os.walk(input_rootpath):
            for file in files:
                if not file.upper().endswith(".IWP"):
                    continue
                filepath = str(os.path.join(root, file))
                micro_vu = MicroVuProgram(filepath)
                print(micro_vu.filepath)
                if micro_vu.has_bring_to_metrology_picture:
                    continue
                if micro_vu.has_been_converted:
                    continue
                self._inject_bring_to_metrology_picture(micro_vu)
                micro_vu.update_instruction_count()
                self._write_file_to_harddrive(micro_vu)


class ProcessorException(Exception):
    pass
