# filename: const.py
# @Time    : 2024/5/8 15:01
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm


DEFAULT_SYMBOL_VALUE_SET = [5, 6, 7, 8, 10]


DEFAULT_CAPABILITY = {
    "textDocument": {
        "synchronization": {
            "dynamicRegistration": False,
            "willSave": True,
            "didSave": True,
            "willSaveWaitUntil": True,
        },
        "publishDiagnostics": {
            "relatedInformation": True,
            "versionSupport": True,
            "codeDescriptionSupport": True,
            "dataSupport": True,
        },
        "diagnostic": {
            "dynamicRegistration": True,
            "relatedDocumentSupport": True,
        },
        "codeAction": {
            "dataSupport": True,
        },
        "documentSymbol": {"symbolKind": {"valueSet": DEFAULT_SYMBOL_VALUE_SET}},
    },
    "workspace": {
        "applyEdit": True,
        "workspaceEdit": {
            "documentChanges": True,
            "resourceOperations": ["create", "rename", "delete"],
        },
        "diagnostics": {
            "refreshSupport": True,
        },
        "fileOperations": {
            "willCreate": True,
            "didCreate": True,
            "didRename": True,
            "didDelete": True,
        },
    },
}
