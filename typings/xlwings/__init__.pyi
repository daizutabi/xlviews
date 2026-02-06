from typing import Any

class Range:
    def __getitem__(self, key: Any) -> Range: ...
    def get_address(
        self,
        row_absolute: bool = True,
        column_absolute: bool = True,
        include_sheetname: bool = False,
        external: bool = False,
    ) -> str: ...
    def offset(self, row_offset: int = 0, column_offset: int = 0) -> Range: ...
