import os
import re
from pathlib import Path
from typing import List

import lib.Utilities
from lib.Utilities import get_utf_encoded_file_lines


class DimensionName:

    def __init__(self, index: int, name: str):
        self.index = index
        self.name = name.upper()


class MicroVuException(Exception):
    pass


class MicroVuProgram:
    _filepath: str
    _has_calculators: bool
    _is_smartprofile: bool
    _op_num: str
    _rev_num: str
    _smartprofile_projectname: str
    _manual_dimension_names: List[DimensionName] = []

    # Static Methods
    @staticmethod
    def get_node(line_text: str, search_value: str) -> str:
        begin_index: int = line_text.upper().find(f"({search_value.upper()}")
        end_index: int = line_text.find(")", begin_index + len(search_value) + 1)
        if end_index == -1:
            end_index = len(line_text)
        return line_text[begin_index:end_index + 1].strip()

    @staticmethod
    def get_node_text(line_text: str, search_value: str, start_delimiter: str, end_delimiter: str = "") -> str:
        if not end_delimiter:
            end_delimiter = start_delimiter
        title_index: int = line_text.upper().find(search_value.upper())
        begin_index: int = line_text.find(start_delimiter, title_index + len(search_value))
        end_index: int = line_text.find(end_delimiter, begin_index + 1)
        if end_index == -1:
            end_index = len(line_text)
        return line_text[begin_index + 1:end_index].strip()

    @staticmethod
    def set_node_text(line_text: str, search_value: str, set_value: str, start_delimiter: str,
                      end_delimiter: str = "") -> str:
        current_value: str = MicroVuProgram.get_node_text(line_text, search_value, start_delimiter, end_delimiter)
        current_node: str = search_value + start_delimiter + current_value + end_delimiter
        new_node: str = search_value + start_delimiter + set_value + end_delimiter
        return line_text.replace(current_node, new_node)

    # Dunder Methods
    def __init__(self, input_filepath: str, op_num: str, rev_num: str, smartprofile_projectname: str):
        self._filepath = input_filepath
        self.file_lines = get_utf_encoded_file_lines(self._filepath)
        self._op_num = op_num.upper()
        self._rev_num = rev_num.upper()
        self._smartprofile_projectname = smartprofile_projectname
        self._postinit()

    # Internal Methods
    def _does_name_already_exist(self, name_to_find: str) -> bool:
        search_text = f"(Name \"{name_to_find}\")"
        return any(line.find(search_text) > 1 for line in self.file_lines)

    def _get_instructions_count(self) -> str:
        return str(len([line for line in self.file_lines if line.find("(Name ") > 1]))

    def _global_replace(self, old_value: str, new_value: str) -> None:
        quoted_old_value = f"\"{old_value}\""
        quoted_new_value = f"\"{new_value}\""
        for i, l in enumerate(self.file_lines):
            if l.find(quoted_old_value) > 0:
                new_line = l.replace(quoted_old_value, quoted_new_value)
                self.file_lines[i] = new_line

    def _postinit(self):
        self._set_smartprofile()
        self._set_has_calculators()

    def _set_has_calculators(self) -> None:
        calcs: list[bool] = [line.startswith("Calc ") for line in self.file_lines]
        self._has_calculators = any(calcs)

    def _set_smartprofile(self) -> None:
        line_idx = self.get_index_containing_text("AutoExpFile")
        existing_export_filepath = str(Path(MicroVuProgram.get_node_text(
                self.file_lines[line_idx], "AutoExpFile", "\""))).upper()
        existing_export_filename = Path(MicroVuProgram.get_node_text(
                self.file_lines[line_idx], "AutoExpFile", "\"")).stem.upper()
        if existing_export_filename == "OUTPUT":
            self._is_smartprofile = True
            return
        if "C:\\MICROVU\\POINTCLOUDS\\" in existing_export_filepath:
            self._is_smartprofile = True
            return
        self._is_smartprofile = False

    # Properties
    @property
    def bring_part_to_metrology_index(self) -> int:
        idx = next(
                (i for i, l in enumerate(self.file_lines)
                 if "Bring part to Metrology.JPG" in l), -1
        )
        if idx > -1:
            return idx

        kill_file_idx = self.kill_file_call_index
        return -1 if kill_file_idx == -1 else kill_file_idx + 3

    @property
    def can_write_to_output_file(self) -> bool:
        return not os.path.exists(self.output_filepath)

    @property
    def comment(self) -> str:
        if comment_idx := self.get_index_containing_text("(Name \"Edited"):
            if " (Txt " in self.file_lines[comment_idx]:
                return MicroVuProgram.get_node_text(self.file_lines[comment_idx], "(Txt ", "\"")
            else:
                return ""
        else:
            return ""

    @comment.setter
    def comment(self, value: str) -> None:
        if line_idx := self.get_index_containing_text("(Name \"Edited"):
            if " (Txt " in self.file_lines[line_idx]:
                updated_comment_line = MicroVuProgram.set_node_text(self.file_lines[line_idx], "(Txt ", value, "\"")
                self.file_lines[line_idx] = updated_comment_line
            else:
                if value.startswith("\r\n"):
                    new_line = f'{self.file_lines[line_idx][:-2]} (Txt \"{value[4:]}\")\n'
                else:
                    new_line = f'{self.file_lines[line_idx][:-2]} (Txt \"{value}\")\n'
                self.file_lines[line_idx] = new_line

    @property
    def dimension_names(self) -> list[DimensionName]:
        if self.is_smartprofile:
            return []
        if self._manual_dimension_names:
            return self._manual_dimension_names
        dimensions: list[DimensionName] = []
        for i, line in enumerate(self.file_lines):
            if line.find("(PropLabels ") > -1:
                if line.startswith("Calc"):
                    continue
                if line.startswith("Prmt"):
                    continue
                old_dimension_name = MicroVuProgram.get_node_text(line, "(Name ", "\"")
                dimension = DimensionName(i, old_dimension_name)
                dimensions.append(dimension)
        return dimensions

    @property
    def export_filepath(self) -> str:
        if self.is_smartprofile:
            return "C:\\TEXT\\OUTPUT.txt"
        if line_idx := self.get_index_containing_text("AutoExpFile"):
            return MicroVuProgram.get_node_text(self.file_lines[line_idx], "AutoExpFile", "\"")
        else:
            return ""

    @export_filepath.setter
    def export_filepath(self, value: str) -> None:
        if self.is_smartprofile:
            value = "C:\\TEXT\\OUTPUT.txt"
        line_idx = self.get_index_containing_text("AutoExpFile")
        if not line_idx:
            return
        line_text = self.file_lines[line_idx]
        updated_line_text = MicroVuProgram.set_node_text(line_text, "(ExpFile ", value, "\"")
        updated_line_text = MicroVuProgram.set_node_text(updated_line_text, "(AutoExpFile ", value, "\"")
        if self.is_smartprofile:
            updated_line_text = updated_line_text.replace("(AutoExpFSApSt DT)", "(AutoExpFSApSt None)")
        else:
            updated_line_text = updated_line_text.replace("(AutoExpFSApSt None)", "(AutoExpFSApSt DT)")
        updated_line_text = updated_line_text.replace("(FldDlm Tab)", "(FldDlm CrLf)")
        self.file_lines[line_idx] = updated_line_text

    @property
    def filename(self) -> str:
        return Path(self._filepath).name

    @property
    def filepath(self) -> str:
        return self._filepath

    @property
    def has_text_kill(self) -> bool:
        return any("C:\\killFile.bat" in line for line in self.file_lines)

    @property
    def has_calculators(self) -> bool:
        return self._has_calculators

    @property
    def has_auto_report(self) -> bool:
        if not (line_idx := self.get_index_containing_text("AutoExpFile")):
            return False
        line = self.file_lines[line_idx]
        if "(AutoRptSortInstructionsByName 0)" not in line:
            return False
        if "(AutoRptTemplateName" not in line:
            return False
        if "(AutoRptAppendDateAndTime" not in line:
            return False
        return "(AutoRptFileName" in line

    @property
    def get_existing_smartprofile_call_index(self) -> int:
        return self.get_index_containing_text("SmartProfile.exe")

    @property
    def instructions_index(self) -> int:
        return next((i for i, l in enumerate(self.file_lines) if l.startswith("Instructions")), 0)

    @property
    def is_smartprofile(self) -> bool:
        return self._is_smartprofile

    @property
    def has_bring_to_metrology_picture(self) -> bool:
        return any("Bring part to Metrology.JPG" in line for line in self.file_lines)

    @property
    def kill_file_call_index(self) -> int:
        return next(
                (i for i, l in enumerate(self.file_lines)
                 if "killFile.bat" in l), -1
        )

    @property
    def last_microvu_system_id(self) -> str:
        last_system_reference_line = [line for line in self.file_lines if line.upper().find("(SYS ") > 1][-1]
        if last_system_reference_line.startswith("Sys 1"):
            return MicroVuProgram.get_node_text(last_system_reference_line, "Sys 1", " ")
        else:
            return MicroVuProgram.get_node_text(last_system_reference_line, "(Sys", " ", ")")

    @property
    def manual_dimension_names(self) -> List[DimensionName]:
        return self._manual_dimension_names

    @manual_dimension_names.setter
    def manual_dimension_names(self, names: List[DimensionName]):
        self._manual_dimension_names = names

    @property
    def op_number(self) -> str:
        return self._op_num

    @property
    def output_directory(self) -> str:
        output_rootpath = lib.Utilities.GetStoredIniValue("Paths", "output_rootpath", "Settings")
        parent_directory = Path(self._filepath).parts[-2]
        return str(Path(output_rootpath, parent_directory))

    @property
    def output_filepath(self) -> str:
        output_directory = self.output_directory
        file_name = Path(self._filepath).name
        return str(Path(output_directory, file_name))

    @property
    def part_number(self) -> str:
        filename = Path(self._filepath).stem
        parts = re.split("[ _]", filename)
        return parts[0]

    @property
    def prompt_insertion_index(self) -> int:
        insert_index: int = self.get_index_containing_text("(Name \"Created")
        if not insert_index or not self.file_lines[insert_index].startswith("Txt"):
            return -1

        temp_idx: int = self.get_index_containing_text("(Name \"Edited")
        if not temp_idx or not self.file_lines[temp_idx].startswith("Txt"):
            return -1
        return max(temp_idx, insert_index)

    @property
    def report_filepath(self) -> str:
        if self.is_smartprofile:
            return ""
        if line_idx := self.get_index_containing_text("AutoRptFileName"):
            return MicroVuProgram.get_node_text(self.file_lines[line_idx], "AutoRptFileName", "\"")
        else:
            return ""

    @report_filepath.setter
    def report_filepath(self, value: str) -> None:
        if self.is_smartprofile:
            value = ""

        if self.has_auto_report:
            line_idx = self.get_index_containing_text("AutoRptFileName")
            if not line_idx:
                return
            line_text = self.file_lines[line_idx]
            if self.is_smartprofile:
                updated_line_text = MicroVuProgram.set_node_text(line_text, "(AutoRptFileName ", "", "\"")
            else:
                updated_line_text = MicroVuProgram.set_node_text(line_text, "(AutoRptFileName ", value, "\"")
            self.file_lines[line_idx] = updated_line_text
        else:
            if line_idx := self.get_index_containing_text("AutoExpFile"):
                line = self.file_lines[line_idx][:-2]
                if "(AutoRptSortInstructionsByName" not in line:
                    line += " (AutoRptSortInstructionsByName 0)"
                else:
                    existing_node = MicroVuProgram.get_node(line, "AutoRptSortInstructionsByName")
                    line = line.replace(existing_node, "(AutoRptSortInstructionsByName 0)")

                if "(AutoRptTemplateName" not in line:
                    line += " (AutoRptTemplateName \"Classic\")"
                else:
                    existing_node = MicroVuProgram.get_node(line, "AutoRptTemplateName")
                    line = line.replace(existing_node, "(AutoRptTemplateName \"Classic\")")

                if "(AutoRptAppendDateAndTime" not in line:
                    line += " (AutoRptAppendDateAndTime 1)"
                else:
                    existing_node = MicroVuProgram.get_node(line, "AutoRptAppendDateAndTime")
                    line = line.replace(existing_node, "(AutoRptAppendDateAndTime 1)")
                line += f" (AutoRptFileName \"{value}\")\n"
                self.file_lines[line_idx] = line

    @property
    def rev_number(self) -> str:
        return self._rev_num

    @property
    def smartprofile_call_insertion_index(self) -> int:
        return len(self.file_lines) if self.is_smartprofile else -1

    @property
    def smartprofile_projectname(self) -> str:
        return self._smartprofile_projectname

    @property
    def view_name(self) -> str:
        rev_begin_idx = 0
        rev_end_idx = 0
        view_name = ""

        filename = Path(self._filepath).stem
        filename_parts = re.split("[ _]", filename)
        count_of_parts = len(filename_parts)
        if count_of_parts == 1:
            return ""
        for x in range(len(filename_parts)):
            if filename_parts[x].upper().startswith("REV"):
                rev_begin_idx = x
                if filename_parts[rev_begin_idx].upper() == "REV":
                    rev_end_idx = rev_begin_idx + 1
                else:
                    rev_end_idx = rev_begin_idx

        if rev_begin_idx == 0:
            for part in range(1, len(filename_parts)):
                view_name += f"{filename_parts[part]} "
        elif rev_begin_idx == 1 and rev_end_idx < count_of_parts - 1:
            for part in range(rev_end_idx, len(filename_parts)):
                view_name += f"{filename_parts[part]} "
        else:
            for part in range(1, rev_end_idx):
                view_name += f"{filename_parts[part]} "
        return view_name.strip()

    # Public Methods
    def delete_line_containing_text(self, text_to_find: str) -> None:
        idx_to_delete = self.get_index_containing_text(text_to_find)
        if idx_to_delete > 0:
            del self.file_lines[idx_to_delete]

    def get_index_containing_text(self, text_to_find: str) -> int:
        return next(
                (i for i, l in enumerate(self.file_lines)
                 if l.upper().find(text_to_find.upper()) > 1), -1
        )

    def insert_line(self, line_index: int, line: str) -> None:
        self.file_lines.insert(line_index, line)

    def update_feature_name(self, line_index: int, feature_name: str) -> None:
        if self._does_name_already_exist(feature_name):
            return
        current_line = self.file_lines[line_index]
        current_name = MicroVuProgram.get_node_text(current_line, "(Name ", "\"")
        new_line = MicroVuProgram.set_node_text(current_line, "(Name ", feature_name, "\"")
        new_line = new_line.replace(f"\"{current_name}\"", f"\"{feature_name}\"")
        self.file_lines[line_index] = new_line
        self._global_replace(f"\"{current_name}\"", f"\"{feature_name}\"")

    def update_instruction_count(self) -> None:
        instruction_count = self._get_instructions_count()
        idx: int = self.get_index_containing_text("AutoExpFile")
        self.file_lines[idx] = MicroVuProgram.set_node_text(
                self.file_lines[idx], "(InsIdx", instruction_count, " ", ")")
        instruction_line_idx = next((i for i, l in enumerate(self.file_lines) if l.startswith("Instructions")), 0)
        self.file_lines[instruction_line_idx] = MicroVuProgram.set_node_text(
                self.file_lines[instruction_line_idx], "Instructions", instruction_count, " ")
