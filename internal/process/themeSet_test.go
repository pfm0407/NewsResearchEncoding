package process

import "testing"

func themeSetTest(t *testing.T) {
	want := "themeSet_Start"
	if got := themeSet(); got != want {
		t.Errorf("themeSet() = %q, want %q", got, want)
	}
}
