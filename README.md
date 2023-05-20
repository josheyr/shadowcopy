# shadowcopy
Shadowcopy allows you to easily create Volume Shadow Copies using WMI API in Golang. Tested on Windows 11

```go
	shadowCopy, err := NewShadowCopy()
	if err != nil {
		t.Fatal(err)
	}

	// log shadow copy ID and volume path (device object)
	t.Log(shadowCopy.ID, shadowCopy.DeviceObject)

	// delete shadow copy
	err = DeleteShadowCopy(shadowCopy.ID)
	if err != nil {
		t.Fatal(err)
	}
  ```
