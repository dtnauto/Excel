package com.example.excel.autogencode

import com.example.excel.usecase52.cutAndFormatString
import java.util.Arrays

fun main(args: Array<String>) {
//    CodeTest().genAllCodeMenu()
    CodeTest().genAllCodeSettingNormal()
}

class CodeTest {
    private val syntaxName = "\$nameSetting$"
    private val syntaxId = "\$nameId$"
    private val syntaxTotalItem = "\$totalItem$"
    private val syntaxValue = "\$nameValue$"
    fun genAllCodeMenu() {
        val listId: List<String> = mutableListOf(
        )
        val listName: List<String> = mutableListOf(
        )
        println("Interface")
        for (i in listId.indices) {
            genInterfaceMenu(listName[i])
            println("")
        }
        println("====================================================================")
        println("====================================================================")
        println("====================================================================")
        println("Code: ")
        for (i in listId.indices) {
            genCodeMenu(listId[i], listName[i])
            println("")
        }
    }

    fun genAllCodeSettingNormal() {
        val listId: List<String> = mutableListOf(
//            "UNKNOWN_ID",
            "FCTA_ON_OFF",
            "FCTA_SENS",
            "RCTA_ON_OFF",
            "LDW_ON_OFF",
            "LDW_WARN",
            "BSM_ON_OFF",
            "BSM_SENS",
            "CBDA_ON_OFF",
            "CBDA_SENS",
            "PEDAL_WARNING",
            "REVERSE_RUNNING_DETECTION",
            "CATEGORY_RISK_DETECTION_RESET"
        )
        val listName: List<String> = mutableListOf(
//            "OncomingVehicleWarningWhenTurning",
            "FrontSideApproachingVehicleDetection",
            "FrontSideApproachingVehicleDetectionWarningTiming",
            "RearSideApproachingVehicleDetectionSetting",
            "LaneDepartureWarningSystem",
            "WarningType",
            "BlindSpotMonitoring",
            "WarningTiming",
            "WarningWhenExitingVehicle",
            "WarningTimingWhenExitingVehicle",
            "PedalOperationWarning",
            "WrongWayDrivingDetection",
            "ResetRiskDetection"
        )
        val listItems = Arrays.asList(
            arrayOf<String>(), // For UNKNOWN_ID, no selections provided
            arrayOf("OFF", "DISPLAY_ONLY", "DISPLAY_AND_ALARM_SOUND"), // For FCTA_ON_OFF
            arrayOf("EARLY", "NORMAL", "SLOW"), // For FCTA_SENS
            arrayOf("OFF", "DISPLAY_ONLY", "DISPLAY_AND_ALARM_SOUND"), // For RCTA_ON_OFF
            arrayOf("ON", "OFF"), // For LDW_ON_OFF
            arrayOf(
                "OFF",
                "WARNING_SOUND",
                "STEERING_VIBRATION",
                "WARNING_SOUND_AND_STEERING_VIBRATION"
            ), // For LDW_WARN
            arrayOf(
                "OFF",
                "DISPLAY_ONLY",
                "DISPLAY_AND_ALARM_SOUND",
                "DISPLAY_AND_STEERING_VIBRATION"
            ), // For BSM_ON_OFF
            arrayOf("EARLY", "NORMAL", "SLOW"), // For BSM_SENS
            arrayOf("OFF", "DISPLAY_ONLY", "DISPLAY_AND_ALARM_SOUND"), // For CBDA_ON_OFF
            arrayOf("EARLY", "NORMAL", "SLOW"), // For CBDA_SENS
            arrayOf("ON", "OFF"), // For PEDAL_WARNING
            arrayOf("ON", "OFF"), // For REVERSE_RUNNING_DETECTION
            arrayOf(
                "RESET_START",
                "RESET_SUCCESS",
                "RESET_FAIL",
                "RESET_IN_PROGRESS"
            ) // For CATEGORY_RISK_DETECTION_RESET
        )
        for (i in listId.indices) {
//            genCodeSettingNormal(listId[i], listName[i], listItems[i])
            genInterfaceSettingNormalMenu(listId[i], listName[i], listItems[i])
            println("")
        }
    }

    fun genInterfaceMenu(name: String) {
        println(
            "boolean get" + name + "Setting(SettingItem setting);"
        )
    }

    fun genInterfaceSettingNormalMenu(id: String, name: String, items: Array<String>) {
        println(
            "boolean get" + name + "Setting(SettingItem setting);"
        )
        for (item in items) {
            println(
                "Result change" + name + "To" + cutAndFormatString(item, format = 2) + "();"
            )
        }
    }


    fun genCodeMenu(id: String?, nameMenu: String?) {
        val outputMenu = """/**
     * @param setting
     * @return
     */
    public boolean get${syntaxName}Setting(SettingItem setting) {
        AFwLog.in(TAG);
        boolean result = false;
        int[] mIds = new int[]{
                SafetySettingIDs.$syntaxId
        };
        SafetySettingInfo[] listSetting = mServiceRepository.getSafetySetting(mIds);
        if (listSetting.length > 0) {
            AFwLog.i(TAG, "listSetting is not empty");
            for (SafetySettingInfo safetySettingInfo : listSetting) {
                result = convert${syntaxName}Setting(safetySettingInfo, setting);
                if (result) {
                    break;
                }
            }
        } else {
            AFwLog.w(TAG, "listSetting is empty");
        }
        AFwLog.i(TAG, String.format("result is %s", result));
        AFwLog.out(TAG);
        return result;
    }

    /**
     * @param info
     * @param settingItem
     * @return
     */
    private boolean convert${syntaxName}Setting(SafetySettingInfo info, SettingItem settingItem) {
        AFwLog.in(TAG);
        int settingId = SafetySettingIDs.$syntaxId;
        if (info == null) {
            AFwLog.w(TAG, "info is null");
            AFwLog.i(TAG, "return false");
            AFwLog.out(TAG);
            return false;
        }
        int infoId = info.getId();
        if (infoId != settingId) {
            AFwLog.w(TAG, String.format("id setting is %s, not %s", infoId, settingId));
            AFwLog.i(TAG, "return false");
            AFwLog.out(TAG);
            return false;
        }

        settingItem.setId(infoId);
        AFwLog.i(TAG, "Found correct setting!");
        validateSupportSafetySettingInfo(info, settingItem);
        AFwLog.i(TAG, "return true");
        AFwLog.out(TAG);
        return true;
    }

    /**
     * @param value
     * @return
     */
    public Result change${syntaxName}Setting(int value) {
        AFwLog.in(TAG);
        int mIds = SafetySettingIDs.$syntaxId;
        Result result = mServiceRepository.changeSafetySetting(
                new int[]{mIds},
                new int[]{value}
        );
        AFwLog.i(TAG, String.format("result is %s", result));
        AFwLog.out(TAG);
        return result;
    }
    """

//        String id = "SAFETY_SETTING_ADS";
//        String nameMenu = "ADS";
        var code = outputMenu.replace(syntaxName, nameMenu!!)
        code = code.replace(syntaxId, id!!)
        println(code)
    }

    fun genCodeSettingNormal(id: String?, name: String?, items: Array<String>) {
        val switchCaseValue = """            case SafetySettingValues.$syntaxValue:
                AFwLog.i(TAG, "value is $syntaxValue");
                break;
"""
        val outputSettingNormal = """     /**
     * @param setting
     * @return
     */
    public boolean get${syntaxName}Setting(SettingItem setting) {
        AFwLog.in(TAG);
        boolean result = false;
        int[] mIds = new int[]{
                SafetySettingIDs.$syntaxId
        };
        SafetySettingInfo[] listSetting = mServiceRepository.getSafetySetting(mIds);
        if (listSetting.length > 0) {
            AFwLog.i(TAG, "listSetting is not empty");
            for (SafetySettingInfo safetySettingInfo : listSetting) {
                result = convert${syntaxName}Setting(safetySettingInfo, setting);
                if (result) {
                    break;
                }
            }
        } else {
            AFwLog.w(TAG, "listSetting is empty");
        }
        AFwLog.i(TAG, String.format("result is %s", result));
        AFwLog.out(TAG);
        return result;
    }

    /**
     * @param info
     * @param settingItem
     * @return
     */
    private boolean convert${syntaxName}Setting(SafetySettingInfo info, SettingItem settingItem) {
        AFwLog.in(TAG);
        int settingId = SafetySettingIDs.$syntaxId;
        int totalItem = $syntaxTotalItem;
        if (info == null) {
            AFwLog.w(TAG, "info is null");
            AFwLog.i(TAG, "return false");
            AFwLog.out(TAG);
            return false;
        }
        int infoId = info.getId();
        if (infoId != settingId) {
            AFwLog.w(TAG, String.format("id setting is %s, not %s", infoId, settingId));
            AFwLog.i(TAG, "return false");
            AFwLog.out(TAG);
            return false;
        }

        AFwLog.i(TAG, "Found correct setting!");
        settingItem.setId(infoId);
        int value = info.getValue();
        AFwLog.i(TAG, "value = " + value);
        settingItem.setValue(value);
        switch (value) {
$syntaxValue            default:
                AFwLog.w(TAG, "value = " + value + "is not designed");
                settingItem.setValue(SettingItem.UNKNOWN_VALUE);
                break;
        }
        validateSupportSafetySettingInfo(info, settingItem);

        // validate list item value
        List<SettingValueItem> listValueItem = new ArrayList<>();
        SafetySettingItemValueInfo[] safetySettingItemValueInfo = info.getMultiValueInfo();
        AFwLog.i(TAG, String.format("safetySettingItemValueInfo have %s item", safetySettingItemValueInfo.length));
        for (SafetySettingItemValueInfo settingItemValueInfo : safetySettingItemValueInfo) {
            if (settingItemValueInfo != null) {
                AFwLog.i(TAG, "settingItemValueInfo is not null");
                SettingValueItem itemValue = new SettingValueItem();
                int valueOfItemValue = settingItemValueInfo.getValue();
                AFwLog.i(TAG, String.format("valueOfItemValue is %s", valueOfItemValue));
                boolean isValidValue = true;
                switch (valueOfItemValue) {
$syntaxValue                    default:
                        AFwLog.w(TAG, "value is invalid");
                        isValidValue = false;
                        break;
                }

                if (isValidValue) {
                    itemValue.setValue(value);
                    validateSupportSafetySettingItemValueInfo(settingItemValueInfo, itemValue);
                    listValueItem.add(itemValue);
                }
            } else {
                AFwLog.w(TAG, "settingItemValueInfo is null");
            }
        }
        if (listValueItem.size() == totalItem) {
            AFwLog.i(TAG, String.format("itemValue have %s item, same as design", totalItem));
        } else {
            AFwLog.w(TAG, String.format("itemValue have %s item, difference from design", listValueItem.size()));
        }
        settingItem.setValueItems(listValueItem);
        AFwLog.i(TAG, "return true");
        AFwLog.out(TAG);
        return true;
    }"""

        /*    */
        /**
         * @param value
         * @return
         *//*
    public Result change${syntaxName}Setting(int value) {
        AFwLog.in(TAG);
        int mIds = SafetySettingIDs.$syntaxId;
        Result result = mServiceRepository.changeSafetySetting(
                new int[]{mIds},
                new int[]{value}
        );
        AFwLog.i(TAG, String.format("result is %s", result));
        AFwLog.out(TAG);
        return result;
    }"""*/


//        String id = "SWITCH_CRUISECONTROL";
//        String nameMenu = "SwitchCruiseControl";
//        String[] items = new String[]{
//                "OFF",
//                "ON"
//        };

        val totalItem = items.size
        val switchValue = StringBuilder()
        for (item in items) {
            switchValue.append(switchCaseValue.replace(syntaxValue, item))
        }
        var code = outputSettingNormal.replace(syntaxName, name!!)
        code = code.replace(syntaxId, id!!)
        code = code.replace(syntaxTotalItem, totalItem.toString())
        code = code.replace(syntaxValue, switchValue.toString())
        code += ""
        println(code)

        for (item in items) {
            val changeFunction = "     /**\n" +
                    "     * @param\n" +
                    "     * @return\n" +
                    "     */\n" +
                    "    public Result change" + name + "To" + cutAndFormatString(
                item,
                format = 2
            ) + "() {\n" +
                    "        AFwLog.in(TAG);\n" +
                    "        int mIds = SafetySettingIDs." + id + ";\n" +
                    "        int mValues = SafetySettingValues." + item + ";\n" +
                    "        Result result = mServiceRepository.changeSafetySetting(\n" +
                    "                new int[]{mIds},\n" +
                    "                new int[]{mValues}\n" +
                    "        );\n" +
                    "        AFwLog.i(TAG, String.format(\"result is %s\", result));\n" +
                    "        AFwLog.out(TAG);\n" +
                    "        return result;\n" +
                    "    }"
            changeFunction.replace(syntaxValue, item)
            println(changeFunction)
        }
    }
}