import json
import os
from dataclasses import dataclass
from pathlib import Path
import adsk.core
import subprocess
import traceback

APP_ID = "UGS_Fusion"
APP_DESC = "Universal G-Code Sender plugin for Fusion"

FUSION_PANEL = "CAMActionPanel"
FUSION_PRODUCT_TYPE = "CAMProductType"

FUSION_BTN_ID = "ugs_fusion_btn_id"
FUSION_BTN_TXT = "Post to UGS"
FUSION_BTN_ICON = "ugs_icon"
FUSION_BTN_TOOL_TIP = "<div style='font-family:\"Calibri\";color:#B33D19; \
    padding-top:-20px;'><span style='font-size:20px;'><b>winder.github.io/ugs_website\
    </b></span></div>Universal Gcode Sender"

GUI_TITLE_TXT = '<a href="http://winder.github.io/ugs_website/">Universal Gcode Sender\
    </a></span><br>A full featured gcode platform used for interfacing \
     with advanced CNC controllers like GRBL and TinyG.'

GUI_FOLDER_ICON = "dir_icon"
GUI_BTN_TEXT = "POST"
GUI_PATHS_LBL = "File paths"
GUI_WIDTH = 580
GUI_HEIGHT = 300

UGS_BINARY_LBL = "UGS binary:"
UGS_BINARY_DEF = "Locate the UGS binary (*.exe,*.jar)"
UGS_BINARY_DLG = "Locate the UGS binary"
UGS_BINARY_FLTR = "*.exe *.jar"

POST_PROCESSOR_LBL = "Post Processor:"
POST_PROCESSOR_DEF = "grbl.cps"
POST_PROCESSOR_DLG = "Select the post processor to use"
POST_PROCESSOR_FLTR = "*.cps"
POST_PROCESSOR_DIR = "Autodesk/Fusion 360 CAM/Posts"

OUTPUT_FOLDER_LBL = "Output Folder:"
OUTPUT_FOLDER_DEF = "Documents/Fusion 360/NC Programs"
OUTPUT_FOLDER_DLG = "Select output folder"

RADIO_GROUP_TXT = "What to Post?"

SHOW_OPERATIONS_DEF = "Setups"
SAVE_INPUT_TXT = "Save settings?"

RESOURCE_DIR = "resources"
SETTINGS_FILE = "settings.json"


# Globals
_handlers = None


def file_path(*args):
    _list = list(args)
    _file = _list.pop()
    _args = reversed(tuple(_list))
    _path = ""
    for _arg in _args:
        _path += str(_arg)
    return str(Path(_path, _file))


def dir_path(*args):
    _path = ""
    for _arg in args:
        _path += str(_arg) + os.sep
    return str(Path(_path)) + os.sep


@dataclass
class Settings:
    ugs_binary: str = UGS_BINARY_DEF
    post_processor: str = POST_PROCESSOR_DEF
    show_operations: str = SHOW_OPERATIONS_DEF
    output_folder: str = dir_path(Path.home(), OUTPUT_FOLDER_DEF)

    def to_json(self):
        return json.dumps(self, default=lambda o: o.__dict__, sort_keys=True, indent=4)


def settings_file():
    app_folder = Path(Path.home(), "." + str.lower(APP_ID))
    if not app_folder.is_dir():
        app_folder.mkdir(True)
    _file = Path(app_folder, SETTINGS_FILE)
    if not _file.is_file():
        _file.write_text(Settings().to_json())
    return _file


def write_settings(settings):
    settings_file().write_text(settings.to_json())


def read_settings():
    settings = json.loads(settings_file().read_text())
    return Settings(**settings)


def export_file(op_name, settings):
    app = adsk.core.Application.get()
    doc = app.activeDocument
    products = doc.products
    product = products.itemByProductType(FUSION_PRODUCT_TYPE)
    fusion_cam = adsk.cam.CAM.cast(product)

    to_post = None

    # Currently doesn't handle duplicate in names
    for setup in fusion_cam.setups:
        if setup.name == op_name:
            to_post = setup
        else:
            for folder in setup.folders:
                if folder.name == op_name:
                    to_post = folder

    for operation in fusion_cam.allOperations:
        if operation.name == op_name:
            to_post = operation

    if to_post == None:
        return 0

    post_processor = file_path(fusion_cam.personalPostFolder, settings.post_processor)
    units = adsk.cam.PostOutputUnitOptions.DocumentUnitsOutput

    # create the postInput object
    post_input = adsk.cam.PostProcessInput.create(
        op_name, post_processor, settings.output_folder, units
    )
    post_input.isOpenInEditor = False
    fusion_cam.postProcess(to_post, post_input)

    # Get the resulting filename
    result_filename = file_path(settings.output_folder, op_name + ".nc")

    # Use subprocess to launch UGS in a new process, check if platform or java
    ugs_binary = Path(settings.ugs_binary)
    if ugs_binary.exists():
        if ugs_binary.suffix == ".exe":
            subprocess.Popen([ugs_binary, "--open", "%s" % result_filename])
        else:
            subprocess.Popen(
                ["java", "-jar", ugs_binary, "--open", "%s" % result_filename]
            )

    return result_filename


# Get the current values of the command inputs.
def get_gui_inputs(gui_inputs):

    settings = Settings(
        ugs_binary=gui_inputs.itemById("binary_txt").text,
        post_processor=gui_inputs.itemById("post_txt").text,
        show_operations=gui_inputs.itemById("operations_group").selectedItem.name,
        output_folder=gui_inputs.itemById("output_txt").text,
    )

    save_settings = gui_inputs.itemById("saveSettings").value
    op_name = None

    # Only attempt to get a value if the user has made a selection
    setup_input = gui_inputs.itemById("setups")
    setup_item = setup_input.selectedItem
    if setup_item:
        setup_name = setup_item.name

    folder_input = gui_inputs.itemById("folders")
    folder_item = folder_input.selectedItem
    if folder_item:
        folder_name = folder_item.name

    operation_input = gui_inputs.itemById("operations")
    operation_item = operation_input.selectedItem
    if operation_item:
        operation_name = operation_item.name

    # Get the name of setup, folder, or operation depending on radio selection
    # This is the operation that will post processed
    if settings.show_operations == "Setups" and setup_item:
        op_name = setup_name
    elif settings.show_operations == "Folders":
        op_name = folder_name
    elif settings.show_operations == "Operations":
        op_name = operation_name

    return op_name, settings, save_settings


# Will update visibility of 3 selection dropdowns based on radio selection
# Also updates radio selection which is only really useful when command is first launched.
def set_dropdown(gui_inputs, show_operations):
    # Get input objects
    setup_input = gui_inputs.itemById("setups")
    folder_input = gui_inputs.itemById("folders")
    radio_group_input = gui_inputs.itemById("operations_group")
    operation_input = gui_inputs.itemById("operations")

    # Set visibility based on appropriate selection from radio list
    if show_operations == "Setups":
        setup_input.isVisible = True
        folder_input.isVisible = False
        operation_input.isVisible = False
        radio_group_input.listItems[0].isSelected = True
    elif show_operations == "Folders":
        setup_input.isVisible = False
        folder_input.isVisible = True
        operation_input.isVisible = False
        radio_group_input.listItems[1].isSelected = True
    elif show_operations == "Operations":
        setup_input.isVisible = False
        folder_input.isVisible = False
        operation_input.isVisible = True
        radio_group_input.listItems[2].isSelected = True
    else:
        # TODO add error check
        return
    return


# Define the event handler for when the command is executed
class GuiBtnHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        try:
            # Get the inputs.
            gui_inputs = args.command.commandInputs
            op_name, settings, save_settings = get_gui_inputs(gui_inputs)

            # Save Settings:
            if save_settings:
                write_settings(settings)

            # Export the file and launch UGS
            file = export_file(op_name, settings)

        except:
            on_exception()


# Define the event handler for when any input changes.
class GuiInputHandler(adsk.core.InputChangedEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):

        try:
            app = adsk.core.Application.get()
            ui = app.userInterface

            # Get inputs and changed inputs
            input_changed = args.input
            gui_inputs = args.inputs

            match input_changed.id:
                case "operations_group":
                    show_operations = input_changed.selectedItem.name
                    set_dropdown(gui_inputs, show_operations)

                # Check for file dialog request
                case "binary_btn":
                    binary_dlg = ui.createFileDialog()
                    binary_dlg.isMultiSelectEnabled = False
                    binary_dlg.title = UGS_BINARY_DEF
                    binary_dlg.filter = UGS_BINARY_FLTR
                    binary_dlg.filterIndex = 1
                    binary_dlg.initialDirectory = str(Path.home())
                    dlgResult = binary_dlg.showOpen()
                    if dlgResult == adsk.core.DialogResults.DialogOK:
                        gui_inputs.itemById("binary_txt").text = file_path(
                            binary_dlg.filename
                        )

                case "post_btn":
                    post_dlg = ui.createFileDialog()
                    post_dlg.isMultiSelectEnabled = False
                    post_dlg.title = POST_PROCESSOR_DLG
                    post_dlg.filter = POST_PROCESSOR_FLTR
                    post_dlg.filterIndex = 1
                    post_dlg.initialDirectory = dir_path(
                        os.getenv("APPDATA"), POST_PROCESSOR_DIR
                    )
                    # mac ~/Library/Application Support/Autodesk/CAM360/libraries/Local
                    dlgResult = post_dlg.showOpen()
                    if dlgResult == adsk.core.DialogResults.DialogOK:
                        gui_inputs.itemById("post_txt").text = file_path(
                            post_dlg.filename
                        )

                case "output_btn":
                    output_dlg = ui.createFolderDialog()
                    output_dlg.title = OUTPUT_FOLDER_DLG
                    output_dlg.initialDirectory = dir_path(
                        Path.home(), OUTPUT_FOLDER_DEF
                    )
                    dlgResult = output_dlg.showDialog()
                    if dlgResult == adsk.core.DialogResults.DialogOK:
                        gui_inputs.itemById("output_txt").text = dir_path(
                            output_dlg.folder
                        )

        except:
            on_exception()


# Define the event handler for when the Add-in is run by the user.
class FusionBtnHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):

        try:
            app = adsk.core.Application.get()
            doc = app.activeDocument
            products = doc.products
            fusion_cam_product = products.itemByProductType(FUSION_PRODUCT_TYPE)
            fusion_cam = adsk.cam.CAM.cast(fusion_cam_product)
            settings = read_settings()

            # Create the gui
            gui = args.command

            # Define the inputs.
            gui_inputs = gui.commandInputs

            # Labels
            gui_inputs.addTextBoxCommandInput("titleTxt", "", GUI_TITLE_TXT, 3, True)
            gui_inputs.addTextBoxCommandInput("group_Text", GUI_PATHS_LBL, "", 1, True)

            # UGS local path and post information
            binary_lbl_input = gui_inputs.addTextBoxCommandInput(
                "binary_lbl", "", UGS_BINARY_LBL, 1, True
            )
            binary_txt_input = gui_inputs.addTextBoxCommandInput(
                "binary_txt", "", settings.ugs_binary, 1, False
            )
            binary_btn_input = gui_inputs.addBoolValueInput(
                "binary_btn", "", False, dir_path(RESOURCE_DIR, GUI_FOLDER_ICON), False
            )
            post_lbl_input = gui_inputs.addTextBoxCommandInput(
                "post_lbl", "", POST_PROCESSOR_LBL, 1, True
            )
            post_txt_input = gui_inputs.addTextBoxCommandInput(
                "post_txt", "", settings.post_processor, 1, False
            )
            post_btn_input = gui_inputs.addBoolValueInput(
                "post_btn", "", False, dir_path(RESOURCE_DIR, GUI_FOLDER_ICON), False
            )
            output_lbl_input = gui_inputs.addTextBoxCommandInput(
                "output_lbl", "", OUTPUT_FOLDER_LBL, 1, True
            )
            output_txt_input = gui_inputs.addTextBoxCommandInput(
                "output_txt", "", settings.output_folder, 1, False
            )
            output_btn_input = gui_inputs.addBoolValueInput(
                "output_btn", "", False, dir_path(RESOURCE_DIR, GUI_FOLDER_ICON), False
            )
            group_table = gui_inputs.addTableCommandInput(
                "group_Tbl", "Table", 0, "7:31:1"
            )
            group_table.tablePresentationStyle = 2
            group_table.addCommandInput(binary_lbl_input, 0, 0)
            group_table.addCommandInput(binary_txt_input, 0, 1)
            group_table.addCommandInput(binary_btn_input, 0, 2)
            group_table.addCommandInput(post_lbl_input, 1, 0)
            group_table.addCommandInput(post_txt_input, 1, 1)
            group_table.addCommandInput(post_btn_input, 1, 2)
            group_table.addCommandInput(output_lbl_input, 2, 0)
            group_table.addCommandInput(output_txt_input, 2, 1)
            group_table.addCommandInput(output_btn_input, 2, 2)

            # What to select from?  Setups, Folders, Operations?
            radio_group_input = gui_inputs.addRadioButtonGroupCommandInput(
                "operations_group", RADIO_GROUP_TXT
            )
            radio_button_items = radio_group_input.listItems
            radio_button_items.add("Setups", False)
            radio_button_items.add("Folders", False)
            radio_button_items.add("Operations", False)

            # Drop down for Setups
            setup_input = gui_inputs.addDropDownCommandInput(
                "setups",
                "Select Setup:",
                adsk.core.DropDownStyles.LabeledIconDropDownStyle,
            )
            # Drop down for Folders
            folder_input = gui_inputs.addDropDownCommandInput(
                "folders",
                "Select Folder:",
                adsk.core.DropDownStyles.LabeledIconDropDownStyle,
            )
            # Drop down for Operations
            operations_input = gui_inputs.addDropDownCommandInput(
                "operations",
                "Select Operation:",
                adsk.core.DropDownStyles.LabeledIconDropDownStyle,
            )
            set_dropdown(gui_inputs, settings.show_operations)

            # Save user settings
            gui_inputs.addBoolValueInput("saveSettings", SAVE_INPUT_TXT, True)
            gui.isExecutedWhenPreEmpted = False

            # Populate values in dropdowns based on current document:
            for setup in fusion_cam.setups:
                setup_input.listItems.add(setup.name, False)
                for folder in setup.folders:
                    folder_input.listItems.add(folder.name, False)
            for operation in fusion_cam.allOperations:
                operations_input.listItems.add(operation.name, False)

            # Defaults for command dialog
            # gui.commandCategoryName = "UGS"
            gui.setDialogInitialSize(GUI_WIDTH, GUI_HEIGHT)
            gui.setDialogMinimumSize(GUI_WIDTH, GUI_HEIGHT)
            gui.okButtonText = GUI_BTN_TEXT

            # Click handler to start app
            onGuiBtn = GuiBtnHandler()
            gui.execute.add(onGuiBtn)
            _handlers.append(onGuiBtn)

            # Input handler to modify gui
            onGuiInput = GuiInputHandler()
            gui.inputChanged.add(onGuiInput)
            _handlers.append(onGuiInput)

        except:
            on_exception()


def run(context):

    # initialize globals
    global _handlers
    app = adsk.core.Application.get()
    ui = app.userInterface
    _handlers = []

    try:

        # Get the existing command definition or create if not exist.
        fusion_btn = ui.commandDefinitions.itemById(FUSION_BTN_ID)
        if not fusion_btn:
            fusion_btn = ui.commandDefinitions.addButtonDefinition(
                FUSION_BTN_ID,
                FUSION_BTN_TXT,
                FUSION_BTN_TOOL_TIP,
                dir_path(RESOURCE_DIR, FUSION_BTN_ICON),
            )

        # Add handler for click event
        on_fusion_btn = FusionBtnHandler()
        fusion_btn.commandCreated.add(on_fusion_btn)
        _handlers.append(on_fusion_btn)

        # Find the "ADD-INS" panel for the solid and the surface workspaces.
        fusion_panel = ui.allToolbarPanels.itemById(FUSION_PANEL)

        # Add a button into both panels.
        fusion_panel.controls.addCommand(fusion_btn, "", True)

        # Pin the button to the panel
        fusion_btn.isPromoted = True

        # fusion_btn.execute()

    except:
        on_exception()


def stop(context):

    try:
        app = adsk.core.Application.get()
        ui = app.userInterface

        # Remove the add-in buttons
        fusion_panel = ui.allToolbarPanels.itemById(FUSION_PANEL)
        fusion_btn = fusion_panel.controls.itemById(FUSION_BTN_ID)
        if fusion_btn:
            fusion_btn.deleteMe()

        # Remove the add-in
        if ui.commandDefinitions.itemById(FUSION_BTN_ID):
            ui.commandDefinitions.itemById(FUSION_BTN_ID).deleteMe()

    except:
        on_exception()


def on_exception():
    app = adsk.core.Application.get()
    ui = app.userInterface
    if ui:
        ui.messageBox("Failed:\n{}".format(traceback.format_exc()))
