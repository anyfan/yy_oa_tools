# UTF-8
#
# For more details about fixed file info 'ffi' see:
# http://msdn.microsoft.com/en-us/library/ms646997.aspx

VSVersionInfo(
    ffi=FixedFileInfo(
        # filevers and prodvers should be always a tuple with four items: (1, 2, 3, 4)
        # Set not needed items to zero 0.
        # 前 16 位是 FileMajorPart 数字。
        # 接下来的 16 位是 FileMinorPart 数字。
        # 第三组 16 位是 FileBuildPart 数字。
        # 最后 16 位是 FilePrivatePart 数字。
        filevers=(${Major}, ${Minor}, ${Revision}, ${Classify}),
        prodvers=(${Major}, ${Minor}, ${Revision}, ${Classify}),
        # Contains a bitmask that specifies the valid bits 'flags'r
        mask=0x3F,
        # Contains a bitmask that specifies the Boolean attributes of the file.
        flags=0x0,
        # The operating system for which this file was designed.
        # 0x4 - NT and there is no need to change it.
        OS=0x4,
        # The general type of file.
        # 0x1 - the file is an application.
        fileType=0x1,
        # The function of the file.
        # 0x0 - the function is not defined for this fileType
        subtype=0x0,
        # Creation date and time stamp.
        date=(0, 0),
    ),
    kids=[
        StringFileInfo(
            [
                StringTable(
                    "000004b0",
                    [
                        StringStruct("CompanyName", "Aerospace Adventure"),
                        StringStruct("FileDescription", "软测办公自动化"),
                        StringStruct("FileVersion", "${git_tag}"),
                        StringStruct("InternalName", "Python Console"),
                        StringStruct(
                            "LegalCopyright",
                            "Copyright © 2024 Aerospace Adventure.",
                        ),
                        StringStruct("OriginalFilename", "oa_tools.exe"),
                        StringStruct("ProductName", "自动化办公工具集"),
                        StringStruct("ProductVersion", "${git_tag}"),
                    ],
                )
            ]
        ),
        VarFileInfo([VarStruct('Translation', [0, 1200])]),
    ],
)