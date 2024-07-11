package main

import (
	"bytes"
	"encoding/json"
	"encoding/xml"
	"gopkg.in/yaml.v3"
	"io"
	"log"
	"net/http"
	"os"
	"time"
)

type Config struct {
	ExchangeURL        string `yaml:"exchange_url"`
	Username           string `yaml:"username"`
	Password           string `yaml:"password"`
	DeviceID           string `yaml:"device_id"`
	CalendarFolderType string `yaml:"calendar_folder_type"`
}

type FolderSyncRequest struct {
	XMLName xml.Name `xml:"FolderSync"`
	SyncKey string   `xml:"SyncKey"`
}

type FolderSyncResponse struct {
	XMLName xml.Name `xml:"FolderSync"`
	Status  int      `xml:"Status"`
	SyncKey string   `xml:"SyncKey"`
	Folders []Folder `xml:"Folders>Folder"`
}

type Folder struct {
	DisplayName string `xml:"DisplayName"`
	Type        string `xml:"Type"`
	FolderID    string `xml:"ServerId"`
}

type Event struct {
	XMLName     xml.Name `xml:"Calendar"`
	Subject     string   `xml:"Subject"`
	StartTime   string   `xml:"Start>DT"`
	EndTime     string   `xml:"End>DT"`
	Description string   `xml:"Body>Content"`
	Location    string   `xml:"Location>DisplayName"`
}

type SyncResponse struct {
	XMLName xml.Name `xml:"Sync"`
	Status  int      `xml:"Status"`
	SyncKey string   `xml:"SyncKey"`
}

func loadConfig(filename string) (*Config, error) {
	data, err := os.ReadFile(filename)
	if err != nil {
		return nil, err
	}

	var config Config
	err = yaml.Unmarshal(data, &config)
	if err != nil {
		return nil, err
	}

	return &config, nil
}

func getFolders(config *Config) ([]Folder, error) {
	syncRequest := FolderSyncRequest{SyncKey: "0"}
	requestBody, err := xml.Marshal(syncRequest)
	if err != nil {
		return nil, err
	}

	req, err := http.NewRequest(
		"POST",
		config.ExchangeURL+"?Cmd=FolderSync&User="+config.Username+"&DeviceId="+config.DeviceID+"&DeviceType=SmartPhone",
		bytes.NewBuffer(requestBody),
	)
	if err != nil {
		return nil, err
	}
	req.SetBasicAuth(config.Username, config.Password)
	req.Header.Set("Content-Type", "application/vnd.ms-sync.wbxml")

	client := &http.Client{}
	resp, err := client.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()

	body, err := io.ReadAll(resp.Body)
	if err != nil {
		return nil, err
	}

	var folderSyncResp FolderSyncResponse
	err = xml.Unmarshal(body, &folderSyncResp)
	if err != nil {
		return nil, err
	}

	return folderSyncResp.Folders, nil
}

func saveFoldersToJSON(folders []Folder, filename string) error {
	data, err := json.MarshalIndent(folders, "", "  ")
	if err != nil {
		return err
	}

	return os.WriteFile(filename, data, 0644)
}

func createEvent(config *Config, calendarID string, event Event) error {
	eventXML, err := xml.Marshal(event)
	if err != nil {
		return err
	}

	req, err := http.NewRequest(
		"POST", config.ExchangeURL+"?Cmd=Sync&User="+config.Username+"&DeviceId="+config.DeviceID+"&DeviceType=SmartPhone",
		bytes.NewBuffer(eventXML),
	)
	if err != nil {
		return err
	}
	req.SetBasicAuth(config.Username, config.Password)
	req.Header.Set("Content-Type", "application/vnd.ms-sync.wbxml")

	client := &http.Client{}
	resp, err := client.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()

	body, err := io.ReadAll(resp.Body)
	if err != nil {
		return err
	}

	var syncResp SyncResponse

	err = xml.Unmarshal(body, &syncResp)
	if err != nil {
		return err
	}

	if syncResp.Status != 1 {
		return err
	}

	return nil
}

func main() {
	config, err := loadConfig("example-config.yaml")
	if err != nil {
		log.Fatalf("Error loading config: %v\n", err)

		return
	}

	folders, err := getFolders(config)
	if err != nil {
		return
	}

	calendars := []Folder{}
	for _, folder := range folders {
		if folder.Type == config.CalendarFolderType {
			calendars = append(calendars, folder)
		}
	}

	err = saveFoldersToJSON(calendars, "calendars.json")
	if err != nil {
		return
	}

	event := Event{
		Subject:     "test",
		StartTime:   time.Now().String(),
		EndTime:     time.Now().Add(time.Hour * 3).String(),
		Description: "djhfbchjbchb",
		Location:    "home",
	}

	if len(calendars) > 0 {
		err = createEvent(config, calendars[0].FolderID, event)
		if err != nil {
			return
		}

	}
}
