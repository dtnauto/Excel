package com.example.excel.autogencode

import java.util.Arrays

fun main() {
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
//        genCodeViewModel(listId[i], listName[i], listItems[i])
        println("")
    }
//    AutoGenCode().genCodeViewModel()
}

class AutoGenCode {
    fun genCodeViewModel(id: String?, name: String?, items: Array<String>) {
        val code = "private SettingItem mFrontApproachingDetection = new SettingItem();"

    }
}