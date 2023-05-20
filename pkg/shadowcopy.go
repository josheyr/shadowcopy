package pkg

import (
	"fmt"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/rs/zerolog/log"
	"golang.org/x/exp/slices"
	"os/exec"
)

type ShadowCopy struct {
	ID           string
	DeviceObject string
}

func NewShadowCopy() (*ShadowCopy, error) {
	err := ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED)
	if err != nil {
		return nil, fmt.Errorf("failed to initialize COM: %w", err)
	}
	defer ole.CoUninitialize()

	wmi, err := connectToWMI()
	if err != nil {
		return nil, err
	}
	defer wmi.Release()

	typeResult, err := getTypeResult(wmi, "Win32_ShadowCopy")
	if err != nil {
		return nil, err
	}
	defer func(typeResult *ole.VARIANT) {
		err := typeResult.Clear()
		if err != nil {
			log.Error().Err(err).Msg("Failed to clear type result.")
		}
	}(typeResult)

	shadowIDsBefore := listShadowCopies(typeResult)

	createResult, err := createShadowCopy(typeResult)
	if err != nil {
		return nil, err
	}
	defer func(createResult *ole.VARIANT) {
		err := createResult.Clear()
		if err != nil {
			log.Error().Err(err).Msg("Failed to clear create result.")
		}
	}(createResult)

	shadowIDsAfter := listShadowCopies(typeResult)

	shadowCopyID := findNewShadowCopyID(shadowIDsBefore, shadowIDsAfter)

	return shadowCopyID, nil
}

func connectToWMI() (*ole.IDispatch, error) {
	unknown, err := oleutil.CreateObject("WbemScripting.SWbemLocator")
	if err != nil {
		return nil, fmt.Errorf("failed to create WMI object: %w", err)
	}

	wmi, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil, fmt.Errorf("failed to query WMI interface: %w", err)
	}

	return wmi, nil
}

func getTypeResult(wmi *ole.IDispatch, typeName string) (*ole.VARIANT, error) {
	serviceRaw, err := oleutil.CallMethod(wmi, "ConnectServer", nil, `root/CIMV2`)
	if err != nil {
		return nil, fmt.Errorf("failed to connect to WMI server: %w", err)
	}
	service := serviceRaw.ToIDispatch()
	defer func(serviceRaw *ole.VARIANT) {
		err := serviceRaw.Clear()
		if err != nil {
			log.Error().Err(err).Msg("Failed to clear service result.")
		}
	}(serviceRaw)

	typeRaw, err := oleutil.CallMethod(service, "Get", typeName)
	if err != nil {
		return nil, fmt.Errorf("failed to get type result: %w", err)
	}

	return typeRaw, nil
}

func createShadowCopy(typeResult *ole.VARIANT) (*ole.VARIANT, error) {
	createResult, err := oleutil.CallMethod(typeResult.ToIDispatch(), "Create", "C:\\", "ClientAccessible")
	if err != nil {
		return nil, fmt.Errorf("failed to create shadow copy: %w", err)
	}

	return createResult, nil
}

func listShadowCopies(typeResult *ole.VARIANT) []*ShadowCopy {
	shadowCopiesRaw, err := oleutil.CallMethod(typeResult.ToIDispatch(), "Instances_", nil)
	if err != nil {
		log.Error().Err(err).Msg("Failed to get shadow copies.")
		return []*ShadowCopy{}
	}

	shadowCopies := shadowCopiesRaw.ToIDispatch()
	defer func(shadowCopiesRaw *ole.VARIANT) {
		err := shadowCopiesRaw.Clear()
		if err != nil {
			log.Error().Err(err).Msg("Failed to clear shadow copies result.")
		}
	}(shadowCopiesRaw)

	var shadowCopyResults []*ShadowCopy
	iterateShadowCopies(shadowCopies, func(item *ole.IDispatch) {
		ID, err := item.GetProperty("ID")
		if err != nil {
			return
		}

		// print VolumeName
		deviceObjectVar, err := item.GetProperty("DeviceObject")
		if err == nil {
			shadowCopyResults = append(shadowCopyResults, &ShadowCopy{ID: ID.ToString(), DeviceObject: deviceObjectVar.ToString()})
		}
	})

	return shadowCopyResults
}

func iterateShadowCopies(shadowCopies *ole.IDispatch, cb func(item *ole.IDispatch)) {
	countVar, err := shadowCopies.GetProperty("Count")
	if err != nil {
		return
	}
	count := int(countVar.Val)
	err = countVar.Clear()
	if err != nil {
		log.Error().Err(err).Msg("Failed to clear count result.")
		return
	}

	for i := 0; i < count; i++ {
		itemRaw, err := oleutil.CallMethod(shadowCopies, "ItemIndex", i)
		if err != nil {
			continue
		}
		item := itemRaw.ToIDispatch()
		//defer itemRaw.Clear()

		cb(item)
	}
}

func findNewShadowCopyID(before, after []*ShadowCopy) *ShadowCopy {
	for _, id := range after {
		if !slices.Contains(before, id) {
			return id
		}
	}

	return nil
}

func DeleteShadowCopy(shadowCopyID string) error {
	// Delete the shadow copy
	_, err := exec.Command("cmd", "/c", "vssadmin", "delete", "shadows", "/shadow="+shadowCopyID, "/Quiet").Output()
	if err != nil {
		return fmt.Errorf("failed to delete shadow copy: %w", err)
	}

	return nil
}
