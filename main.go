package main

import (
	"encoding/csv"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize"
	"golang.org/x/text/encoding/charmap"
)

type statistic struct {
	flag      bool
	firstDate time.Time
	firstDay, bigWeek, allDay,
	unsubscribe, notSend int
}

func parseDate(d string) time.Time {
	var (
		day   int
		month time.Month
		year  int
	)
	ds := strings.Split(d, "-")

	for i := 0; i < len(ds); i++ {

		DateInt, _ := strconv.Atoi(ds[i])

		switch i {
		case 0:
			year = DateInt
		case 1:
			month = time.Month(DateInt)
		default:
			day = DateInt
		}
	}
	return time.Date(year, month, day, 12, 0, 0, 0, time.UTC)
}

func (s *statistic) stat(stat []string, status map[string]string) {

	if !s.flag {
		s.flag = true
		s.firstDate = parseDate(stat[3])

	}

	if (stat[1] == status["open"] || stat[1] == status["openLink"] || stat[1] == status["unsubscribed"]) && stat[2] == stat[3] {
		s.firstDay++
	}

	if (stat[1] == status["open"] || stat[1] == status["openLink"] || stat[1] == status["unsubscribed"]) && parseDate(stat[2]).After(s.firstDate.Add(time.Hour*24*6)) {
		s.bigWeek++
	}

	if stat[1] == status["open"] || stat[1] == status["openLink"] || stat[1] == status["unsubscribed"] {
		s.allDay++
	}

	if stat[1] == status["unsubscribed"] {
		s.unsubscribe++
	}

	if stat[1] == status["hardND"] || stat[1] == status["softND"] {
		s.notSend++
	}
}

func main() {

	var s statistic

	status := map[string]string{
		"unsubscribed": "Отписался от рассылки",
		"hardND":       "Письмо не доставлено (жесткий отказ)",
		"softND":       "Письмо не доставлено (мягкий отказ)",
		"open":         "Письмо открыто",
		"openLink":     "Письмо открыто: переход по ссылке",
		"send":         "Письмо отправлено",
	}

	surchFile, err := filepath.Glob("*.csv")
	if err != nil {
		log.Fatal(err)
	}
	for _, CountCSV := range surchFile {
		s = statistic{false, time.Now(), 0, 0, 0, 0, 0}
		openFile, err := os.Open(CountCSV)
		if err != nil {
			log.Fatal(err)
		}
		defer openFile.Close()
		decodeBin := charmap.Windows1251.NewDecoder().Reader(openFile)
		readCSV := csv.NewReader(decodeBin)
		readCSV.Comma = ';'
		readCSV.LazyQuotes = true

		reader, err := readCSV.ReadAll()
		if err != nil {
			log.Fatal(err)
		}

		var newRead = make([][]string, len(reader)-1)
		copy(newRead, reader[1:])
		for _, list := range newRead {
			s.stat(list, status)
		}
		fmt.Println("Отправлено:", len(newRead))
		fmt.Println("Не доставлено:", len(newRead)-(len(newRead)-s.notSend))
		fmt.Println("Доставлено:", len(newRead)-s.notSend)
		fmt.Println("Открыто всего:", s.allDay)
		fmt.Println("Открыто в 1 день:", s.firstDay)
		fmt.Println("Открыто за неделю:", s.allDay-s.bigWeek-s.firstDay)
		fmt.Println("Открыто больше недели:", s.bigWeek)
		fmt.Println("Отписалось людей:", s.unsubscribe)
		fmt.Println("\n")

		name := strings.Split(CountCSV, ".csv")[0]
		xlsx, err := excelize.OpenFile(name + ".xlsx")
		if err != nil {
			log.Fatal(err)
		}

		xlsx.SetCellValue("Summary", "C7", len(newRead))
		xlsx.SetCellValue("Summary", "D7", len(newRead)-(len(newRead)-s.notSend))
		xlsx.SetCellValue("Summary", "F7", len(newRead)-s.notSend)
		xlsx.SetCellValue("Summary", "H7", s.allDay)
		xlsx.SetCellValue("Summary", "J7", s.firstDay)
		xlsx.SetCellValue("Summary", "L7", s.allDay-s.bigWeek-s.firstDay)
		xlsx.SetCellValue("Summary", "N7", s.bigWeek)
		err1 := xlsx.SaveAs("finish/" + name + ".xlsx")
		if err != nil {
			fmt.Println(err1)
		}
	}
}
