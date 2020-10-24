package xldtparse

import (
	"fmt"
	"testing"
	"time"
)

func TestMMDDYYParse(t *testing.T) {
	s := "01-22-20"
	d, err := parsemmddyyString(s)
	if err != nil {
		t.Error(err)
	}

	fmt.Println(d.String())
}
func TestFloatStringParse(t *testing.T) {
	s := "44015.933692129598"
	d, err := parseFromFloatString(s)
	if err != nil {
		t.Error(err)
	}
	fmt.Println(d.Date())
}
func TestParse(t *testing.T) {
	firstS := "01-22-20"
	d, err := ParseExcelString(firstS)
	if err != nil {
		t.Error(err)
	}
	if d.Day() != 22 {
		t.Error("first day should be 22")
	}
	if d.Year() != 2020 {
		t.Error("first year should be 2020")
	}
	if d.Month() != time.Month(1) {
		t.Error("first month should be 1")
	}

	secondS := "44015.933692129598"
	d2, err := ParseExcelString(secondS)

	if err != nil {
		t.Error(err)
	}
	if d2.Day() != 3 {
		t.Error("first day should be 22")
	}
	if d2.Year() != 2020 {
		t.Error("first year should be 2020")
	}
	if d2.Month() != time.Month(7) {
		t.Error("first month should be 7")
	}
}
