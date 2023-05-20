package pkg

import "testing"

// create shadow copy
func TestCreateDelete(t *testing.T) {
	// create shadow copy
	shadowCopy, err := NewShadowCopy()
	if err != nil {
		t.Fatal(err)
	}

	// log shadow copy ID
	t.Log(shadowCopy.ID)

	// delete shadow copy
	err = DeleteShadowCopy(shadowCopy.ID)
	if err != nil {
		t.Fatal(err)
	}
}
