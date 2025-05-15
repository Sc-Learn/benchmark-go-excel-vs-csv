package main

import (
	"encoding/csv"
	"fmt"
	"os"
	"runtime"
	"time"

	"github.com/xuri/excelize/v2"
)

type Applicant struct {
	Name            string
	Email           string
	CurrentStep     string
	WhatsAppNumber  string
	Experience      string
	PastRole        string
	PastCompany     string
	Campus          string
	MinSalary       string
	CurrentLocation string
	ResumeLink      string
	PortfolioLink   string
	CandidateSource string
}

var headers = []string{
	"Name*", "Email*", "Current Step*", "WhatsApp Number", "Years of Experience",
	"Past Role", "Past Company", "Campus", "Min. Salary", "Current Location",
	"Resume Link", "Portfolio Link", "Candidate Source",
}

func generateDummyApplicants(n int) []Applicant {
	applicants := make([]Applicant, 0, n)
	for i := 0; i < n; i++ {
		applicants = append(applicants, Applicant{
			Name:            fmt.Sprintf("Applicant %d", i),
			Email:           fmt.Sprintf("applicant%d@example.com", i),
			CurrentStep:     "HR Interview",
			WhatsAppNumber:  "6281234567890",
			Experience:      "Junior (1-3 YoE)",
			PastRole:        "Software Engineer",
			PastCompany:     "TechCorp",
			Campus:          "University of Example",
			MinSalary:       "Rp.10,000,000",
			CurrentLocation: "Jakarta",
			ResumeLink:      "https://example.com/resume",
			PortfolioLink:   "https://linkedin.com/in/example",
			CandidateSource: "Website",
		})
	}
	return applicants
}

func exportToExcelStream(filename string, data []Applicant) error {
	f := excelize.NewFile()
	sheet := "Sheet1"

	styleID, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold:  true,
			Color: "#FFFFFF",
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#800080"},
			Pattern: 1,
		},
	})
	if err != nil {
		return err
	}

	streamWriter, err := f.NewStreamWriter(sheet)
	if err != nil {
		return err
	}

	headerRow := make([]interface{}, len(headers))
	for i, h := range headers {
		headerRow[i] = h
		cellName, _ := excelize.CoordinatesToCellName(i+1, 1)
		f.SetCellStyle(sheet, cellName, cellName, styleID)
	}
	if err := streamWriter.SetRow("A1", headerRow); err != nil {
		return err
	}

	for i, a := range data {
		row := []interface{}{
			a.Name, a.Email, a.CurrentStep, a.WhatsAppNumber, a.Experience,
			a.PastRole, a.PastCompany, a.Campus, a.MinSalary,
			a.CurrentLocation, a.ResumeLink, a.PortfolioLink, a.CandidateSource,
		}
		cell, _ := excelize.CoordinatesToCellName(1, i+2)
		if err := streamWriter.SetRow(cell, row); err != nil {
			return err
		}
	}

	if err := streamWriter.Flush(); err != nil {
		return err
	}

	return f.SaveAs(filename)
}

func exportToCSVStream(filename string, data []Applicant) error {
	file, err := os.Create(filename)
	if err != nil {
		return err
	}
	defer file.Close()

	writer := csv.NewWriter(file)
	defer writer.Flush()

	if err := writer.Write(headers); err != nil {
		return err
	}

	for _, a := range data {
		record := []string{
			a.Name, a.Email, a.CurrentStep, a.WhatsAppNumber, a.Experience,
			a.PastRole, a.PastCompany, a.Campus, a.MinSalary,
			a.CurrentLocation, a.ResumeLink, a.PortfolioLink, a.CandidateSource,
		}
		if err := writer.Write(record); err != nil {
			return err
		}
	}

	return nil
}

func printMemoryUsage() {
	var m runtime.MemStats
	runtime.ReadMemStats(&m)
	fmt.Printf("Memory Alloc: %.2f MB\n", float64(m.Alloc)/1024/1024)
	fmt.Printf("Total Alloc : %.2f MB\n", float64(m.TotalAlloc)/1024/1024)
	fmt.Printf("Heap Alloc  : %.2f MB\n", float64(m.HeapAlloc)/1024/1024)
	fmt.Printf("Num GC      : %v\n", m.NumGC)
}

func printFileSize(filename string) {
	info, err := os.Stat(filename)
	if err == nil {
		fmt.Printf("File size: %.2f MB\n", float64(info.Size())/1024.0/1024.0)
	} else {
		fmt.Printf("Error getting file size: %v\n", err)
	}
}

func benchmark(label, filename string, fn func() error) {
	runtime.GC() // clear GC first
	var startMem runtime.MemStats
	runtime.ReadMemStats(&startMem)
	start := time.Now()

	err := fn()

	duration := time.Since(start)

	var endMem runtime.MemStats
	runtime.ReadMemStats(&endMem)

	if err != nil {
		fmt.Printf("[%s] Error: %v\n", label, err)
	} else {
		fmt.Printf("\n[%s] Done in %s\n", label, duration)
		fmt.Printf("Memory delta (Alloc): %.2f MB\n", float64(int64(endMem.Alloc)-int64(startMem.Alloc))/1024.0/1024.0)
		fmt.Printf("Total Alloc Increase: %.2f MB\n", float64(int64(endMem.TotalAlloc)-int64(startMem.TotalAlloc))/1024.0/1024.0)
		fmt.Printf("Heap Alloc Increase : %.2f MB\n", float64(int64(endMem.HeapAlloc)-int64(startMem.HeapAlloc))/1024.0/1024.0)

		printFileSize(filename)
		fmt.Printf("GC Count Increased  : %d\n", endMem.NumGC-startMem.NumGC)
	}
}

func main() {
	count := 100000
	data := generateDummyApplicants(count)

	fmt.Println("Starting export benchmarks with", count, "records...\n")

	benchmark("Export to CSV (Stream)", "benchmark_output.csv", func() error {
		return exportToCSVStream("benchmark_output.csv", data)
	})

	benchmark("Export to Excel (Stream + Style)", "benchmark_output.xlsx", func() error {
		return exportToExcelStream("benchmark_output.xlsx", data)
	})

}
