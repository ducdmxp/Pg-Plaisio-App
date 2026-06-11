using System.Collections.Generic;
using System.Windows.Media;

namespace Convert2DTo3D.Utils
{
    public class Define
    {
        #region SETUP

        public const string DOCKPANEL_ACTIVE_PROJECT_BROWSER = "DOCKPANEL_ACTIVE_PROJECT_BROWSER";
        public const string DOCKPANEL_DEACTIVE_PROJECT_BROWSER = "DOCKPANEL_DEACTIVE_PROJECT_BROWSER";
        public const string DOCKPANEL_COLORING_TAB = "ColoringTab";

        public const string DOCKPANEL_GUID = "950BD36C-9EBC-4585-8C54-76ECC421D406";
        public const string DOCKPANEL_GUID_DISTRIBUTE = "F4B770D9-687D-4670-BC9C-673EF2066A78";

        public const int KEYNOTE_PARAMETER_KEYVALUE = -1140419;
        public const int KEYNOTE_PARAMETER_KEYNOTETEXT = -1140421;
        public static System.Windows.Media.SolidColorBrush BACKGROUND_HIGHTLIGHT = new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 102, 178));
        public static System.Windows.Media.SolidColorBrush BACKGROUND_DEFAULT = new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 255, 255));

        #endregion SETUP

        #region TITLTE

        /// <summary>
        /// Only active color for tab
        /// </summary>
        public const string TITLE_BUTTON_ONLY_ACTIVE_COLOR_FOR_TAB = "TITLE_BUTTON_ONLY_ACTIVE_COLOR_FOR_TAB";

        /// <summary>
        /// My Browser
        /// </summary>
        public const string TITLE_MY_BROWSER = "TITLE_MY_BROWSER";

        /// <summary>
        /// DistributeView
        /// </summary>
        public const string TITLE_TAB_DISTRIBUTE_VIEWS = "TITLE_TAB_DISTRIBUTE_VIEWS";

        /// <summary>
        /// Sheets
        /// </summary>
        public const string TITLE_TAB_SHEETS = "TITLE_TAB_SHEETS";

        /// <summary>
        /// Legends
        /// </summary>
        public const string TITLE_TAB_LEGENDS = "TITLE_TAB_LEGENDS";

        /// <summary>
        /// Schedules/Quantities
        /// </summary>
        public const string TITLE_TAB_SCHEDULES = "TITLE_TAB_SCHEDULES";

        /// <summary>
        /// Views
        /// </summary>
        public const string TITLE_TAB_VIEWS = "TITLE_TAB_VIEWS";

        /// <summary>
        /// DistributeView
        /// </summary>
        public const string TITLE_TAB_DISTRIBUTE = "TITLE_TAB_DISTRIBUTE";

        /// <summary>
        /// Families
        /// </summary>
        public const string TITLE_TAB_FAMILIES = "TITLE_TAB_FAMILIES";

        /// <summary>
        /// All
        /// </summary>
        public const string TITLE_TAB_ALL = "TITLE_TAB_ALL";

        /// <summary>
        /// B1DistributeView
        /// </summary>
        public const string TITLE_DISTRIBUTE_VIEW = "TITLE_DISTRIBUTE_VIEW";

        /// <summary>
        /// Groups
        /// </summary>
        public const string TITLE_TAB_GROUPS = "TITLE_TAB_GROUPS";

        /// <summary>
        /// Detail
        /// </summary>
        public const string TITLE_ITEM_COMPONENT_DETAIL = "TITLE_ITEM_COMPONENT_DETAIL";

        /// <summary>
        /// Model
        /// </summary>
        public const string TITLE_ITEM_COMPONENT_MODEL = "TITLE_ITEM_COMPONENT_MODEL";

        #endregion TITLTE

        #region SORT STRING DEFINE

        public const string SORT_STRING_CONST_1 = "Structural Plans";

        public const string SORT_STRING_CONST_1_JP = "構造伏図";

        public const string SORT_STRING_CONST_2 = "Floor Plans";

        public const string SORT_STRING_CONST_2_JP = "平面図";

        public const string SORT_STRING_CONST_3 = "Ceiling Plans";

        public const string SORT_STRING_CONST_3_JP = "天井伏図";

        public const string SORT_STRING_CONST_4 = "3D Views";

        public const string SORT_STRING_CONST_4_JP = "3D ビュー";

        public const string SORT_STRING_CONST_5 = "Elevations";

        public const string SORT_STRING_CONST_5_JP = "立面図";

        public const string SORT_STRING_CONST_6 = "Sections";

        public const string SORT_STRING_CONST_6_JP = "断面図";

        public const string SORT_STRING_CONST_7 = "Drafting Views";

        public const string SORT_STRING_CONST_7_JP = "製図ビュー";

        public const string SORT_STRING_CONST_8 = "Area Plans";

        public const string SORT_STRING_CONST_8_JP = "エリア プラン";

        #endregion SORT STRING DEFINE

        #region MENU ITEM

        /// <summary>
        /// Back To Default Color
        /// </summary>
        public const string MENU_ITEM_CLEAR_COLOR = "MENU_ITEM_CLEAR_COLOR";

        /// <summary>
        /// Search...
        /// </summary>
        public const string MENU_ITEM_SEARCH = "MENU_ITEM_SEARCH";

        /// <summary>
        /// Expand All
        /// </summary>
        public const string MENU_ITEM_EXPAND_ALL = "MENU_ITEM_EXPAND_ALL";

        /// <summary>
        /// Expand Selected
        /// </summary>
        public const string MENU_ITEM_EXPAND_SELECTED = "MENU_ITEM_EXPAND_SELECTED";

        /// <summary>
        /// Collapse All
        /// </summary>
        public const string MENU_ITEM_COLLAPSE_ALL = "MENU_ITEM_COLLAPSE_ALL";

        /// <summary>
        /// Collapse Selected
        /// </summary>
        public const string MENU_ITEM_COLLAPSE_SELECTED = "MENU_ITEM_COLLAPSE_SELECTED";

        /// <summary>
        /// Go To View
        /// </summary>
        public const string MENU_ITEM_GO_TO_VIEW = "MENU_ITEM_GO_TO_VIEW";

        /// <summary>
        /// Close View
        /// </summary>
        public const string MENU_ITEM_CLOSE_VIEW = "MENU_ITEM_CLOSE_VIEW";

        /// <summary>
        /// Rename View
        /// </summary>
        public const string MENU_ITEM_RENAME_VIEW = "MENU_ITEM_RENAME_VIEW";

        /// <summary>
        /// Reload Family
        /// </summary>
        public const string MENU_ITEM_RELOAD_FAMILY = "MENU_ITEM_RELOAD_FAMILY";

        /// <summary>
        /// Type Properties
        /// </summary>
        public const string MENU_ITEM_TYPE_PROPERTIES = "MENU_ITEM_TYPE_PROPERTIES";

        /// <summary>
        /// Create Instance
        /// </summary>
        public const string MENU_ITEM_CREATE_INSTANCE = "MENU_ITEM_CREATE_INSTANCE";

        /// <summary>
        /// Rename
        /// </summary>
        public const string MENU_ITEM_RENAME = "MENU_ITEM_RENAME";

        /// <summary>
        /// Duplicate View
        /// </summary>
        public const string MENU_ITEM_DUPLICATE_VIEW = "MENU_ITEM_DUPLICATE_VIEW";

        /// <summary>
        /// Duplicate View With Detailing
        /// </summary>
        public const string MENU_ITEM_DUPLICATE_VIEW_WITH_DETAILING = "MENU_ITEM_DUPLICATE_VIEW_WITH_DETAILING";

        /// <summary>
        /// Dulicate View As A Dependent
        /// </summary>
        public const string MENU_ITEM_DUPLICATE_VIEW_AS_A_DEPENDENT = "MENU_ITEM_DUPLICATE_VIEW_AS_A_DEPENDENT";

        /// <summary>
        /// Delete View
        /// </summary>
        public const string MENU_ITEM_DELETE_VIEW = "MENU_ITEM_DELETE_VIEW";

        /// <summary>
        /// Apply View Template
        /// </summary>
        public const string MENU_ITEM_APPLY_VIEW_TEMPLATE = "MENU_ITEM_APPLY_VIEW_TEMPLATE";

        /// <summary>
        /// Delete
        /// </summary>
        public const string MENU_ITEM_DELETE = "MENU_ITEM_DELETE";

        /// <summary>
        /// Duplicate
        /// </summary>
        public const string MENU_ITEM_DUPLICATE = "MENU_ITEM_DUPLICATE";

        /// <summary>
        /// Create Sheet
        /// </summary>
        public const string MENU_ITEM_CREATE_SHEET = "MENU_ITEM_CREATE_SHEET";

        /// <summary>
        /// Change Color BackgroundView
        /// </summary>
        public const string MENU_ITEM_CHANGE_COLOR_FILL_VIEW = "MENU_ITEM_CHANGE_COLOR_FILL_VIEW";

        /// <summary>
        /// Change Color Tab
        /// </summary>
        public const string MENU_ITEM_CHANGE_COLOR_TAB = "MENU_ITEM_CHANGE_COLOR_TAB";

        /// <summary>
        /// Clear Color Tab
        /// </summary>
        public const string MENU_ITEM_CLEAR_COLOR_TAB = "MENU_ITEM_CLEAR_COLOR_TAB";

        /// <summary>
        /// Rename Folder
        /// </summary>
        public const string MENU_ITEM_RENAME_FAVOURITE_FOLDER = "MENU_ITEM_RENAME_FAVOURITE_FOLDER";

        /// <summary>
        /// Delete View In This Folder
        /// </summary>
        public const string MENU_ITEM_DELETE_VIEW_FROM_FAVOURITE_FOLDER = "MENU_ITEM_DELETE_VIEW_FROM_FAVOURITE_FOLDER";

        /// <summary>
        /// Change Color Foreground View
        /// </summary>
        public const string MENU_ITEM_CHANGE_COLOR_VIEW = "MENU_ITEM_CHANGE_COLOR_VIEW";

        /// <summary>
        /// Change Color Border View
        /// </summary>
        public const string MENU_ITEM_CHANGE_COLOR_BORDER_VIEW = "MENU_ITEM_CHANGE_COLOR_BORDER_VIEW";

        /// <summary>
        /// Show View Placement on Sheet Status Icon
        /// </summary>
        public const string MENU_ITEM_SHOW_ICON_VIEW_ON_SHEET = "MENU_ITEM_SHOW_ICON_VIEW_ON_SHEET";

        /// <summary>
        /// Add Favorite Folder
        /// </summary>
        public const string MENU_ITEM_ADD_FAVORITE_FOLDER = "MENU_ITEM_ADD_FAVORITE_FOLDER";

        /// <summary>
        /// Remove
        /// </summary>
        public const string MENU_ITEM_REMOVE_FAVORITE_FOLDER_AND_ITEM = "MENU_ITEM_REMOVE_FAVORITE_FOLDER_AND_ITEM";

        /// <summary>
        /// Add To Favourite Folder
        /// </summary>
        public const string MENU_ITEM_ADD_ITEM_INTO_FAVOURITE_FOLDER = "MENU_ITEM_ADD_ITEM_INTO_FAVOURITE_FOLDER";

        /// <summary>
        /// Find active view
        /// </summary>
        public const string MENU_ITEM_FIND_ACTIVE_VIEW = "MENU_ITEM_FIND_ACTIVE_VIEW";

        /// <summary>
        /// Open sheet
        /// </summary>
        public const string MENU_OPEN_SHEET = "MENU_OPEN_SHEET";

        /// <summary>
        /// Add new profile ...
        /// </summary>
        public const string MENU_ADD_NEW_PROFILE = "MENU_ADD_NEW_PROFILE";

        /// <summary>
        /// Delete selected profile
        /// </summary>
        public const string MENU_DELETE_SELECTED_PROFILE = "MENU_DELETE_SELECTED_PROFILE";

        /// <summary>
        /// Add Color
        /// </summary>
        public const string MENU_ADD_COLOR = "MENU_ADD_COLOR";

        /// <summary>
        /// Remove Color
        /// </summary>
        public const string MENU_REMOVE_COLOR = "MENU_REMOVE_COLOR";

        /// <summary>
        /// Select All Instances
        /// </summary>
        public const string MENU_ITEM_SELECT_ALL_INSTANCES = "MENU_ITEM_SELECT_ALL_INSTANCES";

        /// <summary>
        /// Visible In View
        /// </summary>
        public const string MENU_ITEM_SELECT_VISIBLE_IN_VIEW = "MENU_ITEM_SELECT_VISIBLE_IN_VIEW";

        /// <summary>
        /// Entire Project
        /// </summary>
        public const string MENU_ITEM_SELECT_ENTIRE_PROJECT = "MENU_ITEM_SELECT_ENTIRE_PROJECT";

        #endregion MENU ITEM

        #region MESSAGE

        /// <summary>Can't load the family file. </summary>
        public const string MESS_CAN_NOT_LOAD_FAMILY_FILE = "MESS_CAN_NOT_LOAD_FAMILY_FILE";

        /// <summary>View cannot be placed in sheet. Do you want to continue creating sheets? </summary>
        public const string MESS_VIEW_CAN_NOT_PLACE_ON_SHEET = "MESS_VIEW_CAN_NOT_PLACE_ON_SHEET";

        /// <summary>No matches found </summary>
        public const string MESS_NO_MATCHES_FOUND = "MESS_NO_MATCHES_FOUND";

        /// <summary>Revit has finished searching the project browser. </summary>
        public const string MESS_FINISHED_SEARCHING = "MESS_FINISHED_SEARCHING";

        /// <summary>None </summary>
        public const string MESS_NONE = "MESS_NONE";

        /// <summary>
        /// Add folder before creating a favorites set
        /// </summary>
        public const string MESS_ADD_FOLDER_BEFORE_ADD_SET = "MESS_ADD_FOLDER_BEFORE_ADD_SET";

        /// <summary>
        /// Name cannot contain any of the following characters:\n\\ : { } [ ] | ; < > ? ` ~ \nor any of the non-printable characters.
        /// </summary>
        public const string MESS_INVALID_INPUT = "MESS_INVALID_INPUT";

        /// <summary>
        /// Name is not valid or name is existing!
        /// </summary>
        public const string MESS_INVALID_NAME = "MESS_INVALID_NAME";

        /// <summary>
        /// Name (sheet number) is not valid or name (sheet number) is existing!
        /// </summary>
        public const string MESS_INVALID_NAME_NUMBER_SHEET = "MESS_INVALID_NAME_NUMBER_SHEET";

        /// <summary>
        /// Save name is existing, do you want to override?
        /// </summary>
        public const string MESS_SAVE_DATA_NAME_IS_EXISTING_OVERRIDE = "MESS_SAVE_DATA_NAME_IS_EXISTING_OVERRIDE";

        /// <summary>
        /// Please enter another name!
        /// </summary>
        public const string MESS_SAVE_DATA_ENTER_ANOTHER_NAME = "MESS_SAVE_DATA_ENTER_ANOTHER_NAME";

        /// <summary>
        /// Save name is existing
        /// </summary>
        public const string MESS_SAVE_DATA_NAME_IS_EXISTING = "MESS_SAVE_DATA_NAME_IS_EXISTING";

        /// <summary>
        /// File is in using!
        /// </summary>
        public const string MESS_FILE_IN_USESING = "MESS_FILE_IN_USESING";

        /// <summary>
        /// Command will run after reset application.
        /// </summary>
        public const string MESS_CURRENT_ACTIVE = "MESS_CURRENT_ACTIVE";

        /// <summary>
        /// Command will not run after reset application.
        /// </summary>
        public const string MESS_CURRENT_DEACTIVATE = "MESS_CURRENT_DEACTIVATE";

        /// <summary>
        /// Command is deactivated.
        /// </summary>
        public const string MESS_COMMAND_STATUS_IS_DEACTIVED = "MESS_COMMAND_STATUS_IS_DEACTIVED";

        /// <summary>
        /// Close all selected views before delete
        /// </summary>
        public const string MESS_CLOSE_VIEW_BEFORE_DELETE = "MESS_CLOSE_VIEW_BEFORE_DELETE";

        /// <summary>
        /// Can't find
        /// </summary>
        public const string MESS_CANT_FIND_PATH = "MESS_CANT_FIND_PATH";

        /// <summary>
        /// Please change active view from another project before close this view!
        /// </summary>
        public const string MESS_CHANGE_VIEW_BEFORE_CLOSE = "MESS_CHANGE_VIEW_BEFORE_CLOSE";

        /// <summary>
        /// Please change active view from another project before close this document!
        /// </summary>
        public const string MESS_CHANGE_ACTIVE_VIEW_OTHER_PROJECT_BEFORE_CLOSE = "MESS_CHANGE_ACTIVE_VIEW_OTHER_PROJECT_BEFORE_CLOSE";

        /// <summary>
        /// Favourite folder data has been changed, do you want to save it?
        /// </summary>
        public const string MESS_ASK_SAVE_FAVOURITE_DATA = "MESS_ASK_SAVE_FAVOURITE_DATA";

        /// <summary>
        /// Get data failed!
        /// </summary>
        public const string MESS_GET_DATA_FAILED = "MESS_GET_DATA_FAILED";

        /// <summary>
        /// Turn off the color of the tab
        /// </summary>
        public const string MESS_TURN_OFF_COLOR_TAB = "MESS_TURN_OFF_COLOR_TAB";

        /// <summary>
        /// Turn on the color of the tab
        /// </summary>
        public const string MESS_TURN_ON_COLOR_TAB = "MESS_TURN_ON_COLOR_TAB";

        /// <summary>
        /// Browser settings are saved in this project.
        ///  If you have edited, added, or deleted a favorite, please save the file after clicking the "Save" button.
        /// </summary>
        public const string MESS_ADD_FAVOURITE_OBJECT = "MESS_ADD_FAVOURITE_OBJECT";

        public const string MESS_SAVE_SUCCESS = "MESS_SAVE_SUCCESS";

        /// <summary> Object has been removed from the model. Please reload the panel. </summary>
        public const string MESS_ERROR_ELEMENT_DELETED = "MESS_ERROR_ELEMENT_DELETED";

        public const string MESS_HANDLE_CHANGE_ACTIVEVIEW_ENGLISH = "There is no open view that shows any of the highlighted elements.  Searching through the closed views to find a good view could take a long time.  Continue?";
        public const string MESS_HANDLE_CHANGE_ACTIVEVIEW_JAPANESE = "ハイライト表示された要素すべてを表示するビューが開いていません。適切なビューを探すために閉じたビューから探すにはかなりの時間がかかる可能性があります。続けますか?";

        #endregion MESSAGE

        #region Save favourite data

        public const string FAVOURITE_DATA_NAME_FAMILY = "B_設定保存用ファミリ.rfa";
        public const string FAVOURITE_DATA_NAME_SYMBOL = "B_設定保存用ファミリ";
        public const string FAVOURITE_DATA_NAME_DRAFTING_VIEW = "B_マイブラウザ用設定";
        public const string FAVOURITE_DATA_NAME_DRAFTING_VIEW_FAMILY_TYPE = "設定用";

        public const string FAVOURITE_DATA_PARAMETER_KEY = "Key";
        public const string FAVOURITE_DATA_PARAMETER_NAME = "Name";
        public const string FAVOURITE_DATA_PARAMETER_VALUE = "Value";
        public const string FAVOURITE_DATA_DEFAULT_PARAMETER_KEY = "Convert2DTo3D";
        public const string FAVOURITE_DATA_DEFAULT_TEMP_DATA = "B1FavouriteTempData";

        #endregion Save favourite data

        #region Folder

        public const string FOLDER_FAMILY = "B1Family";
        public const string FOLDER_IMAGE = "Image";
        public const string FOLDER = "Convert2DTo3D";
        public const string FOLDER_COLORING_TAB_PROFILES = "ColoringTabProfiles.json";
        public const string FOLDER_ACTIVE_STATUS = "ActiveStatus.json";

        #endregion Folder

        #region ImageName

        public const string FOLDER_IMAGE_ADD_SAVE_DATA = "AddData";
        public const string FOLDER_IMAGE_ADD_FOLDER = "AddFolder";
        public const string FOLDER_IMAGE_APPLY_FILTER = "ApplyFilterIcon";
        public const string FOLDER_IMAGE_CLEAR_FILTER = "ClearFilterIcon";
        public const string FOLDER_IMAGE_DELETE_SAVE_DATA = "DeleteData";
        public const string FOLDER_IMAGE_EDIT_NAME_SAVE_DATA = "EditData";
        public const string FOLDER_IMAGE_EDIT_NAME = "editname";
        public const string FOLDER_IMAGE_FAMILIES = "Families";
        public const string FOLDER_IMAGE_FAVORITE = "Favorite";
        public const string FOLDER_IMAGE_FOLDER = "folder";
        public const string FOLDER_IMAGE_GROUPS = "Groups";
        public const string FOLDER_IMAGE_GROUPS_DETAIL = "GroupDetail";
        public const string FOLDER_IMAGE_GROUPS_MODEL = "GroupModel";
        public const string FOLDER_IMAGE_GROUPS_ATTACH = "GroupAttach";
        public const string FOLDER_IMAGE_LEGENDS = "Legends";
        public const string FOLDER_IMAGE_RELOAD = "Reload";
        public const string FOLDER_IMAGE_REVIT_LINKS = "Revit Links";
        public const string FOLDER_IMAGE_SAVE_DATA = "SaveData";
        public const string FOLDER_IMAGE_SCHEDULES = "Schedules-Quantities";
        public const string FOLDER_IMAGE_SHEET = "Sheets";
        public const string FOLDER_IMAGE_VIEW_ON_SHEET = "ViewOnSheet";
        public const string FOLDER_IMAGE_VIEW_NOT_ON_SHEET = "ViewNotOnSheet";
        public const string FOLDER_IMAGE_VIEWS = "Views";
        public const string FOLDER_IMAGE_ACTIVE = "Active";
        public const string FOLDER_IMAGE_DEACTIVE = "Deactive";
        public const string FOLDER_IMAGE_CLOSE_ICON = "CloseIcon";
        public const string FOLDER_IMAGE_RFA = "RFAIcon";
        public const string FOLDER_IMAGE_RVT = "RVTIcon";

        #endregion ImageName

        #region TitleControl

        /// <summary>
        /// Warning
        /// </summary>
        public const string CONTROL_TILTLE_WARNING = "CONTROL_TILTLE_WARNING";

        /// <summary>
        /// Annotation symbols
        /// </summary>
        public const string CONTROL_TILTLE_ANNOTATION_SYMBOLS = "CONTROL_TILTLE_ANNOTATION_SYMBOLS";

        /// <summary>
        /// Next
        /// </summary>
        public const string CONTROL_TILTLE_NEXT = "CONTROL_TILTLE_NEXT";

        /// <summary>
        /// Previous
        /// </summary>
        public const string CONTROL_TILTLE_PREVIOUS = "CONTROL_TILTLE_PREVIOUS";

        /// <summary>
        /// Search
        /// </summary>
        public const string CONTROL_TILTLE_SEARCH = "CONTROL_TILTLE_SEARCH";

        /// <summary>
        /// New Folder
        /// </summary>
        public const string CONTROL_TILTLE_NEW_FOLDER = "CONTROL_TILTLE_NEW_FOLDER";

        /// <summary>
        /// New Child Folder
        /// </summary>
        public const string CONTROL_TILTLE_NEW_CHILD_FOLDER = "CONTROL_TILTLE_NEW_CHILD_FOLDER";

        /// <summary>
        /// Cancel
        /// </summary>
        public const string CONTROL_TILTLE_CANCEL = "CONTROL_TILTLE_CANCEL";

        /// <summary>
        /// Sheet number
        /// </summary>
        public const string CONTROL_TILTLE_SHEET_NUMBER = "CONTROL_TILTLE_SHEET_NUMBER";

        /// <summary>
        /// Name
        /// </summary>
        public const string CONTROL_TILTLE_NAME = "CONTROL_TILTLE_NAME";

        #region View Create Sheet

        /// <summary>
        /// Mode
        /// </summary>
        public const string CONTROL_TILTLE_CREATE_SHEET_MODE = "CONTROL_TILTLE_CREATE_SHEET_MODE";

        /// <summary>
        /// Create multies sheet
        /// </summary>
        public const string CONTROL_TILTLE_CREATE_SHEET_MODE_MULTI = "CONTROL_TILTLE_CREATE_SHEET_MODE_MULTI";

        /// <summary>
        /// Create one sheet
        /// </summary>
        public const string CONTROL_TILTLE_CREATE_SHEET_MODE_ONE = "CONTROL_TILTLE_CREATE_SHEET_MODE_ONE";

        /// <summary>
        /// Title Block
        /// </summary>
        public const string CONTROL_TILTLE_CREATE_SHEET_TITLE_BLOCK = "CONTROL_TILTLE_CREATE_SHEET_TITLE_BLOCK";

        /// <summary>
        /// Prefix :
        /// </summary>
        public const string CONTROL_TILTLE_CREATE_SHEET_PREFIX = "CONTROL_TILTLE_CREATE_SHEET_PREFIX";

        /// <summary>
        /// Start number :
        /// </summary>
        public const string CONTROL_TILTLE_CREATE_START_NUMBER = "CONTROL_TILTLE_CREATE_START_NUMBER";

        /// <summary>
        /// Number of digits :
        /// </summary>
        public const string CONTROL_TILTLE_CREATE_NUMBER_DIGITS = "CONTROL_TILTLE_CREATE_NUMBER_DIGITS";

        /// <summary>
        /// Suffix :
        /// </summary>
        public const string CONTROL_TILTLE_CREATE_SUFFIX = "CONTROL_TILTLE_CREATE_SUFFIX";

        /// <summary>
        /// Make the sheet name the same as the view name
        /// </summary>
        public const string CONTROL_TILTLE_CHECKBOX_NAME = "CONTROL_TILTLE_CHECKBOX_NAME";

        /// <summary>
        /// Preview
        /// </summary>
        public const string CONTROL_TILTLE_BTN_PREVIEW = "CONTROL_TILTLE_BTN_PREVIEW";

        #endregion View Create Sheet

        #region View Main Project Browser

        //UI
        /// <summary>
        /// Filter
        /// </summary>
        public const string CONTROL_TILTLE_MAIN_FILTER = "CONTROL_TILTLE_MAIN_FILTER";

        /// <summary>
        /// Detail
        /// </summary>
        public const string CONTROL_TILTLE_MAIN_DETAIL = "CONTROL_TILTLE_MAIN_DETAIL";

        /// <summary>
        /// View Type
        /// </summary>
        public const string CONTROL_TILTLE_MAIN_VIEW_TYPE = "CONTROL_TILTLE_MAIN_VIEW_TYPE";

        /// <summary>
        /// Schedule Type
        /// </summary>
        public const string CONTROL_TILTLE_MAIN_SCHEDULE_TYPE = "CONTROL_TILTLE_MAIN_SCHEDULE_TYPE";

        /// <summary>
        /// View Template
        /// </summary>
        public const string CONTROL_TILTLE_MAIN_VIEW_TEMPLATE = "CONTROL_TILTLE_MAIN_VIEW_TEMPLATE";

        /// <summary>
        /// View Placed On Sheet
        /// </summary>
        public const string CONTROL_TILTLE_MAIN_VIEW_PLACED_ON_SHEET = "CONTROL_TILTLE_MAIN_VIEW_PLACED_ON_SHEET";

        /// <summary>
        /// View is not placed on sheet
        /// </summary>
        public const string CONTROL_TILTLE_MAIN_VIEW_PLACED_NOT_ON_SHEET = "CONTROL_TILTLE_MAIN_VIEW_PLACED_NOT_ON_SHEET";

        /// <summary>
        /// Apply
        /// </summary>
        public const string CONTROL_TILTLE_MAIN_APPLY = "CONTROL_TILTLE_MAIN_APPLY";

        /// <summary>
        /// Clear
        /// </summary>
        public const string CONTROL_TILTLE_MAIN_CLEAR = "CONTROL_TILTLE_MAIN_CLEAR";

        /// <summary>
        /// View Name
        /// </summary>
        public const string CONTROL_TILTLE_MAIN_VIEW_NAME = "CONTROL_TILTLE_MAIN_VIEW_NAME";

        /// <summary>
        /// --select target--
        /// </summary>
        public const string CONTROL_TILTLE_MAIN_SELECT_TARGET = "CONTROL_TILTLE_MAIN_SELECT_TARGET";

        /// <summary>
        /// Target
        /// </summary>
        public const string CONTROL_TILTLE_MAIN_TARGET = "CONTROL_TILTLE_MAIN_TARGET";

        /// <summary>
        /// Reload
        /// </summary>
        public const string CONTROL_TILTLE_MAIN_RELOAD = "CONTROL_TILTLE_MAIN_RELOAD";

        /// <summary>
        /// Favourite
        /// </summary>
        public const string CONTROL_TILTLE_MAIN_FAVOURITE = "CONTROL_TILTLE_MAIN_FAVOURITE";

        #endregion View Main Project Browser

        #region View preview control

        /// <summary>
        /// Close
        /// </summary>
        public const string CONTROL_TILTLE_PREVIEW_CONTROL_CLOSE = "CONTROL_TILTLE_PREVIEW_CONTROL_CLOSE";

        #endregion View preview control

        #region View Save Favourite Folder

        /// <summary>
        /// Save Name
        /// </summary>
        public const string CONTROL_TILTLE_SAVE_FAVOURITE_SAVE_NAME = "CONTROL_TILTLE_SAVE_FAVOURITE_SAVE_NAME";

        #endregion View Save Favourite Folder

        #region View Search

        /// <summary>
        /// Find view :
        /// </summary>
        public const string CONTROL_TILTLE_FIND_VIEW = "CONTROL_TILTLE_FIND_VIEW";

        /// <summary>
        /// Match case
        /// </summary>
        public const string CONTROL_TILTLE_MATCH_CASE = "CONTROL_TILTLE_MATCH_CASE";

        #endregion View Search

        #region Dialog Input

        /// <summary>
        /// Field is required!
        /// </summary>
        public const string CONTROL_TILTLE_VALIDATION_WARNING = "CONTROL_TILTLE_VALIDATION_WARNING";

        #endregion Dialog Input

        /// <summary>
        /// Set color for all tabs
        /// </summary>
        public const string CONTROL_TILTLE_COLORING_TAB_SET_ALL = "CONTROL_TILTLE_COLORING_TAB_SET_ALL";

        /// <summary>
        /// Tab Style
        /// </summary>
        public const string CONTROL_TILTLE_COLORING_TAB_TAB_STYLE = "CONTROL_TILTLE_COLORING_TAB_TAB_STYLE";

        /// <summary>
        /// Color Profiles
        /// </summary>
        public const string CONTROL_TILTLE_COLORING_TAB_PROFILES = "CONTROL_TILTLE_COLORING_TAB_PROFILES";

        /// <summary>
        /// Add new profile
        /// </summary>
        public const string CONTROL_TILTLE_COLORING_TAB_ADD_NEW_PROFILE = "CONTROL_TILTLE_COLORING_TAB_ADD_NEW_PROFILE";

        /// <summary>
        /// Rename selected profile
        /// </summary>
        public const string CONTROL_TILTLE_COLORING_TAB_RENAME_SELECTED_PROFILE = "CONTROL_TILTLE_COLORING_TAB_RENAME_SELECTED_PROFILE";

        /// <summary>
        /// Save data
        /// </summary>
        public const string CONTROL_TILTLE_COLORING_TAB_SAVE_DATA = "CONTROL_TILTLE_COLORING_TAB_SAVE_DATA";

        /// <summary>
        /// Set for seleted project
        /// </summary>
        public const string CONTROL_TILTLE_COLORING_TAB_SET_PROJECT = "CONTROL_TILTLE_COLORING_TAB_SET_PROJECT";

        /// <summary>
        /// Text Color
        /// </summary>
        public const string CONTROL_TILTLE_COLORING_TAB_TEXT_COLOR = "CONTROL_TILTLE_COLORING_TAB_TEXT_COLOR";

        /// <summary>
        /// Auto set text color for tab"
        /// </summary>
        public const string CONTROL_TILTLE_COLORING_TAB_AUTO_SET_TEXTBLOCK = "CONTROL_TILTLE_COLORING_TAB_AUTO_SET_TEXTBLOCK";

        /// <summary>
        /// Top Bar - Light
        /// </summary>
        public const string CONTROL_TILTLE_COLORING_TAB_TOP_LIGHT = "CONTROL_TILTLE_COLORING_TAB_TOP_LIGHT";

        /// <summary>
        /// Top Bar - Medium
        /// </summary>
        public const string CONTROL_TILTLE_COLORING_TAB_TOP_MEDIUM = "CONTROL_TILTLE_COLORING_TAB_TOP_MEDIUM";

        /// <summary>
        /// Top Bar - Heavy
        /// </summary>
        public const string CONTROL_TILTLE_COLORING_TAB_TOP_HEAVY = "CONTROL_TILTLE_COLORING_TAB_TOP_HEAVY";

        /// <summary>
        /// Top Bar - Heavier
        /// </summary>
        public const string CONTROL_TILTLE_COLORING_TAB_TOP_HEAVIER = "CONTROL_TILTLE_COLORING_TAB_TOP_HEAVIER";

        /// <summary>
        /// Border - Light
        /// </summary>
        public const string CONTROL_TILTLE_COLORING_TAB_BORDER_LIGHT = "CONTROL_TILTLE_COLORING_TAB_BORDER_LIGHT";

        /// <summary>
        /// Border - Medium
        /// </summary>
        public const string CONTROL_TILTLE_COLORING_TAB_BORDER_MEDIUM = "CONTROL_TILTLE_COLORING_TAB_BORDER_MEDIUM";

        /// <summary>
        /// Border - Heavy
        /// </summary>
        public const string CONTROL_TILTLE_COLORING_TAB_BORDER_HEAVY = "CONTROL_TILTLE_COLORING_TAB_BORDER_HEAVY";

        /// <summary>
        /// Border - Heavier
        /// </summary>
        public const string CONTROL_TILTLE_COLORING_TAB_BORDER_HEAVIER = "CONTROL_TILTLE_COLORING_TAB_BORDER_HEAVIER";

        /// <summary>
        /// Background Fill
        /// </summary>
        public const string CONTROL_TILTLE_COLORING_TAB_BACKGROUND_FILL = "CONTROL_TILTLE_COLORING_TAB_BACKGROUND_FILL";

        /// <summary>
        /// Filter Mode
        /// </summary>
        public const string CONTROL_TILTLE_FILLTER_MODE_HEADER = "CONTROL_TILTLE_FILLTER_MODE_HEADER";

        #endregion TitleControl

        #region TooltipControl

        /// <summary> Add Favourite Folder </summary>
        public const string CONTROL_TOOLTIP_ADD_FOLDER = "CONTROL_TOOLTIP_ADD_FOLDER";

        /// <summary> Add</summary>
        public const string CONTROL_TOOLTIP_ADD_OBJECT = "CONTROL_TOOLTIP_ADD_OBJECT";

        /// <summary> Rename</summary>
        public const string CONTROL_TOOLTIP_RENAME_OBJECT = "CONTROL_TOOLTIP_RENAME_OBJECT";

        /// <summary> Delete</summary>
        public const string CONTROL_TOOLTIP_DELETE_OBJECT = "CONTROL_TOOLTIP_DELETE_OBJECT";

        /// <summary> Save</summary>
        public const string CONTROL_TOOLTIP_SAVE_OBJECT = "CONTROL_TOOLTIP_SAVE_OBJECT";

        #endregion TooltipControl

        #region Right Click Tab Header

        /// <summary> 3D View </summary>
        public const string TITLE_RIGHTCLICK_3D_VIEW = "TITLE_RIGHTCLICK_3D_VIEW";

        /// <summary> Default 3D View </summary>
        public const string TITLE_RIGHTCLICK_DEFAULT_3D_VIEW = "TITLE_RIGHTCLICK_DEFAULT_3D_VIEW";

        /// <summary> Camera </summary>
        public const string TITLE_RIGHTCLICK_CAMERA = "TITLE_RIGHTCLICK_CAMERA";

        /// <summary> Walkthrough </summary>
        public const string TITLE_RIGHTCLICK_WALKTHROUGH = "TITLE_RIGHTCLICK_WALKTHROUGH";

        /// <summary> Section </summary>
        public const string TITLE_RIGHTCLICK_SECTION = "TITLE_RIGHTCLICK_SECTION";

        /// <summary> Callout </summary>
        public const string TITLE_RIGHTCLICK_CALLOUT = "TITLE_RIGHTCLICK_CALLOUT";

        /// <summary> Rectangle </summary>
        public const string TITLE_RIGHTCLICK_RECTANGLE = "TITLE_RIGHTCLICK_RECTANGLE";

        /// <summary> Sketch </summary>
        public const string TITLE_RIGHTCLICK_SKETCH = "TITLE_RIGHTCLICK_SKETCH";

        /// <summary> Plan View </summary>
        public const string TITLE_RIGHTCLICK_PLAN_VIEW = "TITLE_RIGHTCLICK_PLAN_VIEW";

        /// <summary> Floor Plan </summary>
        public const string TITLE_RIGHTCLICK_FLOOR_PLAN = "TITLE_RIGHTCLICK_FLOOR_PLAN";

        /// <summary> Reflected Ceiling Plan </summary>
        public const string TITLE_RIGHTCLICK_REFLECTED_CEILING_PLAN = "TITLE_RIGHTCLICK_REFLECTED_CEILING_PLAN";

        /// <summary> Structural Plan </summary>
        public const string TITLE_RIGHTCLICK_STRUCTURAL_PLAN = "TITLE_RIGHTCLICK_STRUCTURAL_PLAN";

        /// <summary> Plan Region </summary>
        public const string TITLE_RIGHTCLICK_PLAN_REGION = "TITLE_RIGHTCLICK_PLAN_REGION";

        /// <summary> Area Plan </summary>
        public const string TITLE_RIGHTCLICK_AREA_PLAN = "TITLE_RIGHTCLICK_AREA_PLAN";

        /// <summary> Elevation </summary>
        public const string TITLE_RIGHTCLICK_ELEVATION = "TITLE_RIGHTCLICK_ELEVATION";

        /// <summary> Framing Elevation </summary>
        public const string TITLE_RIGHTCLICK_FRAMING_ELEVATION = "TITLE_RIGHTCLICK_FRAMING_ELEVATION";

        /// <summary> Drafting View </summary>
        public const string TITLE_RIGHTCLICK_DRAFTING_VIEW = "TITLE_RIGHTCLICK_DRAFTING_VIEW";

        /// <summary> Legend </summary>
        public const string TITLE_RIGHTCLICK_LEGEND = "TITLE_RIGHTCLICK_LEGEND";

        /// <summary> Keynote Legend </summary>
        public const string TITLE_RIGHTCLICK_KEYNOTELEGEND = "TITLE_RIGHTCLICK_KEYNOTELEGEND";

        /// <summary> Schedule/Quantities </summary>
        public const string TITLE_RIGHTCLICK_SCHEDULE_QUANTITIES = "TITLE_RIGHTCLICK_SCHEDULE_QUANTITIES";

        /// <summary> Graphical Column Schedule </summary>
        public const string TITLE_RIGHTCLICK_GRAPHICAL_COLUMN_SCHEDULE = "TITLE_RIGHTCLICK_GRAPHICAL_COLUMN_SCHEDULE";

        /// <summary> Material Takeoff </summary>
        public const string TITLE_RIGHTCLICK_MATERIAL_TAKEOFF = "TITLE_RIGHTCLICK_MATERIAL_TAKEOFF";

        /// <summary> Sheet List </summary>
        public const string TITLE_RIGHTCLICK_SHEET_LIST = "TITLE_RIGHTCLICK_SHEET_LIST";

        /// <summary> Note Block </summary>
        public const string TITLE_RIGHTCLICK_NOTE_BLOCK = "TITLE_RIGHTCLICK_NOTE_BLOCK";

        /// <summary> View List </summary>
        public const string TITLE_RIGHTCLICK_VIEW_LIST = "TITLE_RIGHTCLICK_VIEW_LIST";

        /// <summary> Sheet </summary>
        public const string TITLE_RIGHTCLICK_SHEET = "TITLE_RIGHTCLICK_SHEET";

        #endregion Right Click Tab Header

        public static Dictionary<string, string> NameViewType = new Dictionary<string, string>()
        {
            {"Ceiling Plan","天井伏図"},
            {"Floor Plan","平面図"},
            {"Engineering Plan","構造伏図"},
            {"Structural Plan","構造伏図"},
            {"Area Plan","エリア プラン"},
            {"Elevation" ,"立面図"},
            {"Section" ,"断面図"},
            {"3D View","3D ビュー"},
            {"Schedule","集計"},
            {"Drafting","製図ビュー"},
            {"Sheet" ,"シート"},
            {"GraphicalColumnSchedule" ,"柱リスト図"},
            {"Walkthrough","ウォークスルー"},
            {"Rendering" ,"xレンダリング"},
            {"Legend","凡例"},
        };
    }
}