// Package to provide a parsing method for excel strings.
// Intended to be used with the excelize (https://github.com/360EntSecGroup-Skylar/excelize) library as
// many raw string formats are found in the current parsing functions.
package xldtparse

import (
	"errors"
	"strconv"
	"strings"
	"time"
)

const secondsPerDay = 86400

// Standard Year to add if len 2 year text is less than 70
const pre1970YearAdd = 2000

// Standard Year to add if len 2 year text is greater than or equal to 70
const post1970YearAdd = 1990

// ErrNoValidParsingMethods is returns when ParseExcelString is called on a string
// in which no available methods are able to handle it.
var ErrNoValidParsingMethods = errors.New("no suitable parsing methods for current string")

var errParseDatePartParseFail = errors.New("error parsing date part with given constraints")
var errSplitNotRequiredLength = errors.New("split not required length")

func parseDatePart(datePart string, reqLen int, minVal int, maxVal int)(int, error){
	if len(datePart) != reqLen{
		return 0, errors.New("date part not required length")
	}
	if parsedDatePart, err := strconv.ParseInt(datePart, 10, 64); err == nil && int(parsedDatePart) >= minVal && int(parsedDatePart) <= maxVal{
		return int(parsedDatePart), nil
	}
	return 0, errParseDatePartParseFail
}

func parsemmddyyyySlashSeperated(s string)(time.Time, error){
	strSplit := strings.Split(s, "/")
	if len(strSplit) != 3 {
		return time.Time{}, errSplitNotRequiredLength
	}
	yearVal, err := parseDatePart(strSplit[2], 4, 1899, 9999)
	if err != nil{
		return time.Time{}, err
	}
	monthVal, err := parseDatePart(strSplit[0], 2, 1, 12)
	if err != nil{
		return time.Time{}, err
	}
	dayVal, err := parseDatePart(strSplit[1], 2, 1, 31)
	if err != nil{
		return time.Time{}, err
	}

	var dateYear int
	if yearVal < 70 {
		dateYear = yearVal + pre1970YearAdd
	} else {
		dateYear = yearVal + post1970YearAdd
	}
	timeVal := time.Date(dateYear, time.Month(monthVal), int(dayVal), 0, 0, 0, 0, time.UTC)
	return timeVal, nil
}

// ParseExcelDateString attempts to parse a string of the format
// mm-dd-yy into a valid Go time.Time value.
func parsemmddyyString(s string) (time.Time, error) {
	strSplit := strings.Split(s, "-")
	if len(strSplit) != 3 {
		return time.Time{}, errSplitNotRequiredLength
	}
	yearVal, err := parseDatePart(strSplit[2], 2, 0, 99)
	if err != nil{
		return time.Time{}, err
	}
	monthVal, err := parseDatePart(strSplit[0], 2, 1, 12)
	if err != nil{
		return time.Time{}, err
	}
	dayVal, err := parseDatePart(strSplit[1], 2, 1, 31)
	if err != nil{
		return time.Time{}, err
	}

	var dateYear int
	if yearVal < 70 {
		dateYear = yearVal + pre1970YearAdd
	} else {
		dateYear = yearVal + post1970YearAdd
	}
	timeVal := time.Date(dateYear, time.Month(monthVal), int(dayVal), 0, 0, 0, 0, time.UTC)
	return timeVal, nil
}

var excelEpoch = time.Date(1899, time.December, 30, 0, 0, 0, 0, time.UTC)

// parseFromFloatString attempts to parse a string in the excel decimal
// format where the number is days since Dec 30, 1899.
func parseFromFloatString(s string) (time.Time, error) {
	days, err := strconv.ParseFloat(s, 64)
	if err != nil {
		return time.Time{}, err
	}
	return excelEpoch.Add(time.Second * time.Duration(days*secondsPerDay)), nil
}

func parseYYYYdashMMdashDDspTimeStringDatePart(s string)(time.Time, error){
	strSplit := strings.Split(s, "-")

	if len(strSplit) != 3 {
		return time.Time{}, errSplitNotRequiredLength
	}

	yearVal, yearErr := parseDatePart(strSplit[0], 4, 1899, 9999)
	monthVal, monthErr := parseDatePart(strSplit[1], 2, 1, 12)
	dayVal, dayErr := parseDatePart(strSplit[2], 2, 1, 31)

	if yearErr != nil || monthErr != nil || dayErr != nil {
		return time.Time{}, errParseDatePartParseFail
	}

	return time.Date(yearVal, time.Month(monthVal), dayVal, 0, 0, 0, 0, time.UTC), nil
}

func parseYYYdashMMdashDDspTimeStringTimePart(s string)(time.Duration, error){
	strSplit := strings.Split(s, ":")

	if len(strSplit) != 3 {
		return time.Second, errSplitNotRequiredLength
	}

	hourVal, hourErr := parseDatePart(strSplit[0], 2, 0, 23)
	minuteVal, minuteErr := parseDatePart(strSplit[0], 2, 0, 59)
	secondVal, secondErr := parseDatePart(strSplit[0], 2, 0, 59)

	if hourErr != nil || minuteErr != nil || secondErr != nil {
		return time.Second, errParseDatePartParseFail
	}

	return time.Hour * time.Duration(hourVal) + time.Minute * time.Duration(minuteVal) + time.Second * time.Duration(secondVal), nil
}
func parseYYYYdashMMdashDDspTimeString(s string)(time.Time, error){
	dateTimeSplit := strings.Split(s, " ")

	if len(dateTimeSplit) != 2{
		return time.Time{}, errSplitNotRequiredLength
	}
	datePart, dateErr := parseYYYYdashMMdashDDspTimeStringDatePart(dateTimeSplit[0])
	timePart, timeErr := parseYYYdashMMdashDDspTimeStringTimePart(dateTimeSplit[1])

	if timeErr != nil || dateErr != nil {
		return time.Time{}, errParseDatePartParseFail
	}
	return datePart.Add(timePart), nil
}

// ParseExcelString attempts to parse an excel string value
// to a Go time.Time value.
func ParseExcelString(s string) (time.Time, error) {
	if parseVal, err := parseFromFloatString(s); err == nil{
		return parseVal, nil
	}
	if parseVal, err := parsemmddyyString(s); err == nil{
		return parseVal, nil
	}
	if parseVal, err := parsemmddyyyySlashSeperated(s); err == nil{
		return parseVal, nil
	}
	if parseVal, err := parseYYYYdashMMdashDDspTimeString(s); err == nil{
		return parseVal, nil
	}
	return time.Time{}, ErrNoValidParsingMethods
}
