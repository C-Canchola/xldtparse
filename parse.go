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
const Pre1970YearAdd = 2000

// Standard Year to add if len 2 year text is greater than or equal to 70
const Post1970YearAdd = 1990

// ErrNoValidParsingMethods is returns when ParseExcelString is called on a string
// in which no available methods are able to handle it.
var ErrNoValidParsingMethods = errors.New("no suitable parsing methods for current string")

var errNotLenThree = errors.New("length not three")
var errSplitTextNotLenTwo = errors.New("split item length not two")
var errParseMonthValueInvalid = errors.New("parse month value invalid")
var errParseDayValueInvalid = errors.New("parsed day value invalid")
var errParseYearValueInvalid = errors.New("parsed year value invalid")

// ParseExcelString attempts to parse an excel string value
// to a Go time.Time value.
func ParseExcelString(s string) (time.Time, error) {
	numericParse, err := parseFromFloatString(s)
	if err == nil {
		return numericParse, nil
	}
	return parsemmddyyString(s)
}

// ParseExcelDateString attempts to parse a string of the format
// mm-dd-yy into a valid Go time.Time value.
func parsemmddyyString(s string) (time.Time, error) {
	strSplit := strings.Split(s, "-")
	if len(strSplit) != 3 {
		return time.Time{}, errNotLenThree
	}
	for _, s := range strSplit {
		if len(s) != 2 {
			return time.Time{}, errSplitTextNotLenTwo
		}
	}
	monthVal, err := strconv.ParseInt(strSplit[0], 10, 64)
	if err != nil {
		return time.Time{}, err
	}
	if monthVal > 12 || monthVal < 1 {
		return time.Time{}, errParseMonthValueInvalid
	}
	dayVal, err := strconv.ParseInt(strSplit[1], 10, 64)
	if err != nil {
		return time.Time{}, err
	}
	if dayVal < 1 || dayVal > 31 {
		return time.Time{}, errParseDayValueInvalid
	}
	yearVal, err := strconv.ParseInt(strSplit[2], 10, 64)
	if err != nil {
		return time.Time{}, err
	}
	if yearVal < 1 {
		return time.Time{}, errParseYearValueInvalid
	}
	var dateYear int
	if yearVal < 70 {
		dateYear = int(yearVal + Pre1970YearAdd)
	} else {
		dateYear = int(yearVal + Post1970YearAdd)
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
