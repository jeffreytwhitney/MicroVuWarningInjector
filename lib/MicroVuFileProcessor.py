import contextlib
import os
import re
import string
from abc import ABCMeta, abstractmethod
from datetime import datetime
from pathlib import Path
from typing import List

import lib.Utilities
from lib import Utilities
from lib.MicroVuProgram import MicroVuProgram, MicroVuException, DimensionName
from lib.Utilities import get_unencoded_file_lines, get_utf_encoded_file_lines, get_filepath_by_name


def _convert_number_to_letter(num_value: str) -> str:
    charstr = ' ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    chars = list(charstr)
    return chars[int(num_value)]


def _collapse_dot_dimension_parts(parts: list[str]) -> list[str]:
    if any(part for part in parts if part == "."):
        if parts[0] == ".":
            return []
        if parts[-1] == ".":
            return []
        for i, part in enumerate(parts):
            if parts[i] == ".":
                parts[i - 1] = f"{parts[i - 1]}.{parts[i + 1]}"
                parts.remove(parts[i + 1])
                parts.remove(parts[i])
    return parts


def _containsNumber(line):
    res = False
    if " " in line:
        return False
    if "_" in line:
        return False
    with contextlib.suppress(ValueError):
        for val in line.split():
            if float(val.strip(string.punctuation)):
                res = True
                break
    return res


def _get_dimension_parts(dimension_name: str) -> list[str]:
    dimension_name = dimension_name.replace("INSP", "").replace("ITEM", "")
    parts = re.split("(\\d+)", dimension_name)
    while "" in parts:
        parts.remove("")
    while " " in parts:
        parts.remove(" ")
    while "_" in parts:
        parts.remove("_")
    while "-" in parts:
        parts.remove("-")
    for i, part in enumerate(parts):
        parts[i] = part.strip().replace("_", "").replace("-", "")
    return _collapse_dot_dimension_parts(parts)


class Processor(metaclass=ABCMeta):
    _dimension_root: str
    _hand_edit_dimension_names: bool
    _microvu_programs: List[MicroVuProgram] = []

    @abstractmethod
    def process_files(self) -> None:
        pass

    def __init__(self, user_initials: str):
        self.user_initials = user_initials
        self._dimension_root: str = lib.Utilities.GetStoredIniValue("GlobalSettings", "dimension_root", "Settings")
        hand_edit_setting_value = lib.Utilities.GetStoredIniValue("GlobalSettings", "hand_edit_dimension_names", "Settings")
        self._hand_edit_dimension_names = hand_edit_setting_value == "True"

    @staticmethod
    def parse_dimension_name(dimension_name: str, dimension_root: str) -> str:
        dimension_name = dimension_name.upper()
        if dimension_name.startswith("#"):
            dimension_name = dimension_name[1:]
        if _containsNumber(dimension_name):
            return f"{dimension_root}{dimension_name}"
        if dimension_name.isalpha():
            return dimension_name

        parts = _get_dimension_parts(dimension_name)

        if not parts:
            return dimension_name
        if not _containsNumber(parts[0]):
            return dimension_name

        num_parts = len(parts)
        match num_parts:

            case 1:
                if _containsNumber(parts[0]):
                    return f"{dimension_root}{parts[0]}"
                else:
                    return dimension_name

            case 2:
                if _containsNumber(parts[0]) and parts[1].isnumeric():
                    return f"{dimension_root}{parts[0]}{_convert_number_to_letter(parts[1])}"
                if _containsNumber(parts[0]) and parts[1].isalpha():
                    if len(parts[1]) > 1:
                        return dimension_name
                    if parts[1] == "X":
                        return dimension_name
                    return f"{dimension_root}{parts[0]}{parts[1]}"
                return dimension_name

            case 3:
                if parts[2] != "X":
                    return dimension_name
                if _containsNumber(parts[0]) and parts[1].isnumeric():
                    return f"{dimension_root}{parts[0]}{_convert_number_to_letter(parts[1])}"
                if _containsNumber(parts[0]) and parts[1].isalpha():
                    if len(parts[1]) > 1:
                        return dimension_name
                    if parts[1] == "X":
                        return dimension_name
                    return f"{dimension_root}{parts[0]}{parts[1]}"
                return dimension_name

            case 4:
                if parts[-1] != "X":
                    return dimension_name
                if not parts[0].isnumeric():
                    return dimension_name
                if parts[1].isalpha():
                    return dimension_name
                if parts[2].isnumeric():
                    return f"{dimension_root}{parts[0]}{parts[1]}{_convert_number_to_letter(parts[2])}"
                if parts[2].isalpha() and len(parts[2]) == 1:
                    return f"{dimension_root}{parts[0]}{parts[1]}{parts[2]}"
                return dimension_name

            case _:
                return dimension_name

    @property
    def allow_deletion_of_old_program(self) -> bool:
        return Utilities.GetStoredIniValue("GlobalSettings", "allow_delete", "Settings") == "True"

    @property
    def disable_on_convert(self) -> bool:
        return Utilities.GetStoredIniValue("GlobalSettings", "disable_on_convert", "Settings") == "True"

    @property
    def micro_vu_programs(self) -> list[MicroVuProgram]:
        return self._microvu_programs

    def add_micro_vu_program(self, micro_vu: MicroVuProgram):
        self._microvu_programs.append(micro_vu)

    def add_micro_vu_programs(self, micro_vus: list[MicroVuProgram]):
        for micro_vu in micro_vus:
            self._microvu_programs.append(micro_vu)


class CoonRapidsProcessor(Processor):

    def _delete_old_prompts(self, micro_vu):
        micro_vu.delete_line_containing_text("Name \"PT #\"")
        micro_vu.delete_line_containing_text("Name \"Employee #\"")
        micro_vu.delete_line_containing_text("Name \"Machine #\"")
        micro_vu.delete_line_containing_text("Name \"PT#\"")
        micro_vu.delete_line_containing_text("Name \"Employee#\"")
        micro_vu.delete_line_containing_text("Name \"Machine#\"")
        micro_vu.delete_line_containing_text("Name \"Run-Setup\"")
        micro_vu.delete_line_containing_text("Name \"Job #\"")
        micro_vu.delete_line_containing_text("Name \"Job#\"")
        micro_vu.delete_line_containing_text("Name \"PT\"")
        micro_vu.delete_line_containing_text("Name \"REV LETTER\"")
        micro_vu.delete_line_containing_text("Name \"OPERATION\"")
        micro_vu.delete_line_containing_text("Name \"EMPLOYEE\"")
        micro_vu.delete_line_containing_text("Name \"JOB\"")
        micro_vu.delete_line_containing_text("Name \"MACHINE\"")
        micro_vu.delete_line_containing_text("Name \"IN PROCESS\"")
        micro_vu.delete_line_containing_text("Name \"SEQUENCE\"")
        micro_vu.delete_line_containing_text("Name \"SPFILENAME\"")

    def _disable_dimensions(self, micro_vu):
        for i, line in enumerate(micro_vu.file_lines):
            if "AutoExpFile" in line:
                line = line.replace("(AutoRpt 1)", "(AutoRpt 0)")
                line = line.replace("(AutoConf 1)", "(AutoConf 0)")
                micro_vu.file_lines[i] = line
                continue
            if "(Name " not in line:
                continue
            if "C:\\killFile.bat" in line:
                continue
            if "Bring part to Metrology.JPG" in line:
                continue
            if "(DontMeasure)" in line:
                continue
            micro_vu.file_lines[i] = f"{line[:-1]} (DontMeasure){line[-1]}"

    def _get_new_prompts(self, micro_vu: MicroVuProgram) -> list[str]:
        pattern = 'sp_prompt_text.txt' if micro_vu.is_smartprofile else 'prompt_text.txt'

        prompt_file = get_filepath_by_name(pattern)
        if not prompt_file:
            return []

        prompt_lines = get_utf_encoded_file_lines(prompt_file)
        return prompt_lines or []

    def _inject_bring_to_metrology_picture(self, micro_vu: MicroVuProgram) -> None:

        if micro_vu.kill_file_call_index == -1:
            return

        bring_to_met_pic_idx = micro_vu.bring_part_to_metrology_index

        bring_to_met_text_filepath = get_filepath_by_name('BringPartToMetrology_text.txt')
        if not bring_to_met_text_filepath:
            raise ProcessorException("Can't find 'BringPartToMetrology_text' file.")

        lines = get_unencoded_file_lines(bring_to_met_text_filepath)
        if not lines:
            raise ProcessorException("Can't find 'BringPartToMetrology_text' file.")

        micro_vu.insert_line(bring_to_met_pic_idx, lines[3])
        micro_vu.insert_line(bring_to_met_pic_idx, lines[2])
        micro_vu.insert_line(bring_to_met_pic_idx, lines[1])

    def _inject_kill_file_call(self, micro_vu: MicroVuProgram) -> None:

        if micro_vu.instructions_index == -1:
            return

        if any("killFile.bat" in line for line in micro_vu.file_lines):
            self._replace_kill_file_call(micro_vu)
            return

        text_kill_index = micro_vu.instructions_index + 1

        textkill_filepath = get_filepath_by_name('TextKill_text.txt')
        if not textkill_filepath:
            raise ProcessorException("Can't find 'TextKill_text' file.")

        lines = get_unencoded_file_lines(textkill_filepath)
        if not lines:
            raise ProcessorException("Can't find 'TextKill_text' file.")

        micro_vu.insert_line(text_kill_index, lines[3])
        micro_vu.insert_line(text_kill_index, lines[2])
        micro_vu.insert_line(text_kill_index, lines[1])

    def _inject_smart_profile_call(self, micro_vu: MicroVuProgram) -> None:

        if not micro_vu.is_smartprofile:
            return

        existing_smartprofile_call_index = micro_vu.get_existing_smartprofile_call_index
        if existing_smartprofile_call_index > -1:
            micro_vu.file_lines[existing_smartprofile_call_index] = micro_vu.file_lines[
                existing_smartprofile_call_index].replace("UniversalSmartProfile.py", "OneFactorySP.py")
            return

        smartprofile_call_insertion_index = micro_vu.smartprofile_call_insertion_index
        if smartprofile_call_insertion_index == -1:
            return

        microvu_system_id: str = micro_vu.last_microvu_system_id
        if not microvu_system_id:
            return

        smartprofile_filepath = get_filepath_by_name('CallSmartProfile_text.txt')
        if not smartprofile_filepath:
            raise ProcessorException("Can't find 'CallSmartProfile_text' file.")

        prompt_lines = get_unencoded_file_lines(smartprofile_filepath)
        if not prompt_lines:
            raise ProcessorException("Can't find 'CallSmartProfile_text' file.")

        smartprofile_line = prompt_lines[1]
        smartprofile_script_path = lib.Utilities.GetStoredIniValue("Paths", "smart_profile_script_filepath", "Settings")
        smartprofile_exe_path = lib.Utilities.GetStoredIniValue("Paths", "smart_profile_exe_filepath", "Settings")
        smartprofile_line = smartprofile_line.replace("<?SYS>", microvu_system_id)
        smartprofile_line = smartprofile_line.replace("<?EXE>", smartprofile_exe_path)
        smartprofile_line = smartprofile_line.replace("<?SCR>", smartprofile_script_path)
        micro_vu.file_lines.append(smartprofile_line)
        micro_vu.file_lines.append(prompt_lines[2])
        micro_vu.file_lines.append(prompt_lines[3])

    def _replace_dimension_names(self, micro_vu: MicroVuProgram) -> None:
        if micro_vu.is_smartprofile:
            return
        dimension_names: list[DimensionName] = micro_vu.dimension_names
        for dimension_name in dimension_names:
            if self._hand_edit_dimension_names:
                new_dimension_name = dimension_name.name
            else:
                new_dimension_name = Processor.parse_dimension_name(dimension_name.name, self._dimension_root)
            micro_vu.update_feature_name(dimension_name.index, new_dimension_name)

    def _replace_export_filepath(self, micro_vu: MicroVuProgram) -> None:
        if micro_vu.is_smartprofile:
            micro_vu.export_filepath = "C:\\TEXT\\OUTPUT.txt"
            return
        part_rev = f"REV{micro_vu.rev_number}"
        export_filepath: str = "C:\\Users\\Public\\CURL\\in\\"
        export_filepath += micro_vu.part_number
        export_filepath += f"_OP{micro_vu.op_number}"
        view_name = str(micro_vu.view_name)
        if view_name != "":
            export_filepath += f"_{micro_vu.view_name}"
        export_filepath += f"_{part_rev}"
        export_filepath += "_.csv"
        micro_vu.export_filepath = export_filepath

    def _replace_kill_file_call(self, micro_vu: MicroVuProgram):
        killfile_index = next(
                (i for i, l in enumerate(micro_vu.file_lines)
                 if "killFile.bat" in l), -1
        )
        if killfile_index == -1:
            return
        current_line = micro_vu.file_lines[killfile_index]
        killfile_node = MicroVuProgram.get_node(current_line, "CmdText")
        new_line = current_line.replace(killfile_node, "(CmdText \"\"\"C:\\killFile.bat\"\"\")")
        micro_vu.file_lines[killfile_index] = new_line

    def _replace_prompt_section(self, micro_vu: MicroVuProgram) -> None:
        insert_index = micro_vu.prompt_insertion_index
        if insert_index == -1:
            raise ProcessorException("There is either no 'Edited By' or 'Created By' line. Cannot process file.")
        insert_index += 1

        prompt_lines = self._get_new_prompts(micro_vu)
        if not prompt_lines:
            raise ProcessorException("Can't find 'prompt_text' file.")

        self._delete_old_prompts(micro_vu)

        for line in prompt_lines:
            if line.find("(Name \"IN PROCESS\")") > 0:
                micro_vu.insert_line(insert_index, line)
                continue
            if line.find("(Name \"MACHINE\")") > 0:
                micro_vu.insert_line(insert_index, line)
                continue
            if line.find("(Name \"JOB\")") > 0:
                micro_vu.insert_line(insert_index, line)
                continue
            if line.find("(Name \"EMPLOYEE\")") > 0:
                micro_vu.insert_line(insert_index, line)
                continue
            if line.find("(Name \"OPERATION\")") > 0:
                line = line.replace("<O>", str(micro_vu.op_number))
                micro_vu.insert_line(insert_index, line)
                continue
            if line.find("(Name \"REV LETTER\")") > 0:
                line = line.replace("<R>", str(micro_vu.rev_number))
                micro_vu.insert_line(insert_index, line)
                continue
            if line.find("(Name \"PT\")") > 0:
                line = line.replace("<P>", str(micro_vu.part_number))
                micro_vu.insert_line(insert_index, line)
                continue
            if line.find("(Name \"SEQUENCE\")") > 0:
                micro_vu.insert_line(insert_index, line)
                continue
            if line.find("(Name \"SPFILENAME\")") > 0:
                smartprofile_projectname = micro_vu.smartprofile_projectname
                line = line.replace("<SPF>", smartprofile_projectname)
                micro_vu.insert_line(insert_index, line)
                continue
        return

    def _replace_report_filepath(self, micro_vu: MicroVuProgram) -> None:
        if micro_vu.is_smartprofile:
            micro_vu.report_filepath = ""
            return
        view_name = micro_vu.view_name
        part_rev = f"REV{micro_vu.rev_number}"
        report_filepath: str = Utilities.GetStoredIniValue("Paths", "reporting_root_path", "Settings")
        report_filepath += micro_vu.part_number
        report_filepath += f"_OP{micro_vu.op_number}"
        if len(view_name) > 0:
            report_filepath += f"_{view_name}"
        report_filepath += f"_{part_rev}_.pdf"
        micro_vu.report_filepath = report_filepath

    def _update_comments(self, micro_vu: MicroVuProgram) -> None:
        date_text = datetime.now().strftime("%m/%d/%Y")
        new_comment = f"\\r\\nConverted program to work with 1Factory. {self.user_initials} {date_text}."
        current_comment = micro_vu.comment
        current_comment += new_comment
        micro_vu.comment = current_comment

    def _write_file_to_harddrive(self, micro_vu: MicroVuProgram) -> None:
        if os.path.exists(micro_vu.output_filepath):
            file_name = Path(micro_vu.output_filepath).name
            dir_name = os.path.dirname(micro_vu.output_filepath)
            raise ProcessorException(
                    f"File '{file_name}' already exists in output directory '{dir_name}'."
            )
        os.makedirs(micro_vu.output_directory, exist_ok=True)
        with open(micro_vu.output_filepath, 'w+', encoding='utf-16-le', newline='\r\n') as f:
            for line in micro_vu.file_lines:
                f.write(f"{line}")

    def process_files(self) -> None:
        try:
            for micro_vu in self.micro_vu_programs:
                self._replace_export_filepath(micro_vu)
                self._replace_report_filepath(micro_vu)
                if not micro_vu.is_smartprofile:
                    self._replace_dimension_names(micro_vu)
                else:
                    self._inject_smart_profile_call(micro_vu)
                self._replace_prompt_section(micro_vu)
                if not micro_vu.has_text_kill:
                    self._inject_kill_file_call(micro_vu)
                if self.disable_on_convert and not micro_vu.has_bring_to_metrology_picture:
                    self._inject_bring_to_metrology_picture(micro_vu)
                self._update_comments(micro_vu)
                if self.disable_on_convert:
                    self._disable_dimensions(micro_vu)
                micro_vu.update_instruction_count()
                self._write_file_to_harddrive(micro_vu)
                if self.allow_deletion_of_old_program:
                    os.remove(micro_vu.filepath)
        except Exception as e:
            raise ProcessorException(e.args[0]) from e


class AnokaProcessor(CoonRapidsProcessor):
    def __init__(self, user_initials: str):
        super().__init__(user_initials)

    def _replace_prompt_section(self, micro_vu: MicroVuProgram) -> None:
        try:
            super()._replace_prompt_section(micro_vu)
        except ProcessorException:
            self._anoka_replace_prompt_section(micro_vu)
        return

    def _anoka_replace_prompt_section(self, micro_vu: MicroVuProgram) -> None:
        insert_idx: int = micro_vu.get_index_containing_text("(Name \"START")
        if insert_idx == 0:
            raise ProcessorException(f"Can't determine where to put the prompts. Cannot process file {micro_vu.filename}.")

        start_idx: int = micro_vu.get_index_containing_text("AutoExpFile")
        for idx in range(insert_idx, start_idx, -1):
            if micro_vu.file_lines[idx].startswith("Prmt"):
                del micro_vu.file_lines[idx]

        insert_index: int = micro_vu.get_index_containing_text("(Name \"START")

        prompt_lines = self._get_new_prompts(micro_vu)
        if not prompt_lines:
            raise ProcessorException("Can't find 'prompt_text' file.")

        for line in prompt_lines:
            if line.find("(Name \"IN PROCESS\")") > 0:
                micro_vu.insert_line(insert_index, line)
                continue
            if line.find("(Name \"MACHINE\")") > 0:
                micro_vu.insert_line(insert_index, line)
                continue
            if line.find("(Name \"JOB\")") > 0:
                micro_vu.insert_line(insert_index, line)
                continue
            if line.find("(Name \"EMPLOYEE\")") > 0:
                micro_vu.insert_line(insert_index, line)
                continue
            if line.find("(Name \"OPERATION\")") > 0:
                line = line.replace("<O>", str(micro_vu.op_number))
                micro_vu.insert_line(insert_index, line)
                continue
            if line.find("(Name \"REV LETTER\")") > 0:
                line = line.replace("<R>", str(micro_vu.rev_number))
                micro_vu.insert_line(insert_index, line)
                continue
            if line.find("(Name \"PT\")") > 0:
                line = line.replace("<P>", str(micro_vu.part_number))
                micro_vu.insert_line(insert_index, line)
                continue
            if line.find("(Name \"SEQUENCE\")") > 0:
                micro_vu.insert_line(insert_index, line)
                continue
            if line.find("(Name \"SPFILENAME\")") > 0:
                try:
                    smartprofile_projectname = micro_vu.smartprofile_projectname
                except MicroVuException as e:
                    raise ProcessorException(e.args[0]) from e
                line = line.replace("<SPF>", smartprofile_projectname)
                micro_vu.insert_line(insert_index, line)
                continue
        return


def get_processor(user_initials: str):
    if Utilities.GetStoredIniValue("Location", "Site", "Settings") == "CoonRapids":
        processor = CoonRapidsProcessor(user_initials)
    else:
        processor = AnokaProcessor(user_initials)

    processor.micro_vu_programs.clear()
    return processor


class ProcessorException(Exception):
    pass
