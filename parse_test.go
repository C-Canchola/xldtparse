package xldtparse

import (
	"fmt"
	"testing"
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
