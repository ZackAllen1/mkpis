package ui

import (
	"fmt"
	"log"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/common-nighthawk/go-figure"
	"github.com/davidscholberg/go-durationfmt"
	"github.com/jmartin82/mkpis/pkg/vcs"
	"github.com/olekukonko/tablewriter"
)

var f = excelize.NewFile()

func AvgDurationFormater(d time.Duration) string {
	t, err := durationfmt.Format(d, "AVG: %dd %hh %mm")
	if err != nil {
		return "ERROR"
	}
	return t
}

func DurationFormater(d time.Duration) string {

	if d.Microseconds() == 0 {
		return "--"
	}

	t, err := durationfmt.Format(d, "%hh %mm")
	if err != nil {
		return "ERROR"
	}
	return t
}

type CmdUI struct {
	client       vcs.Client
	owner        string
	repo         string
	develBranch  string
	masterBranch string
}

func NewCmdUI(client vcs.Client, owner, repo, develBranch, masterBranch string) *CmdUI {
	return &CmdUI{
		client:       client,
		owner:        owner,
		repo:         repo,
		develBranch:  develBranch,
		masterBranch: masterBranch,
	}
}

func (u CmdUI) Render(from, to time.Time) error {
	rfb, err := u.getFeatureBranchReport(from, to)
	if err != nil {
		return err
	}
	rrb, err := u.getReleaseBranchReport(from, to)
	if err != nil {
		return err
	}

	myFigure := figure.NewColorFigure("Printing the reports...", "standard", "white", true)
	myFigure.Blink(1000, 300, 300)

	fmt.Println("\033[2J") //clean previous ouput
	u.PrintPageHeader(from, to)
	u.PrintRepotHeader("Feature Branch Report")
	fmt.Println(rfb)
	u.PrintRepotHeader("Release Branch Report")
	fmt.Println(rrb)

	if err := f.SaveAs("../../GithubBranchReport.xlsx"); err != nil {
		log.Fatal(err)
	}
	return nil
}

func (u CmdUI) PrintRepotHeader(text string) {
	figure.NewColorFigure(text, "small", "green", true).Print()
	fmt.Println("")
}

func (u CmdUI) PrintPageHeader(from time.Time, to time.Time) {
	figure.NewColorFigure("MKPIS", "standard", "red", true).Print()
	fLayout := "2006-02-01"
	fmt.Printf("\n Repo: %s/%s (%s-%s)", u.owner, u.repo, from.Format(fLayout), to.Format(fLayout))
	fmt.Println("")
}

func (u CmdUI) getFeatureBranchReport(from, to time.Time) (string, error) {
	prs, err := u.client.GetMergedPRList(u.owner, u.repo, from, to, u.develBranch)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error gathering information: %s", err.Error())
		return "", err
	}

	infoSheet := "Info"
	f.NewSheet(infoSheet)
	f.SetColWidth(infoSheet, "A", "I", 20)
	f.SetCellValue(infoSheet, "A1", "Owner")
	f.SetCellValue(infoSheet, "B1", "Repo")
	f.SetCellValue(infoSheet, "C1", "From")
	f.SetCellValue(infoSheet, "D1", "To")
	f.SetCellValue(infoSheet, "E1", "FeatureBranchName")
	f.SetCellValue(infoSheet, "F1", "MasterBranchName")

	f.SetCellValue(infoSheet, "A2", u.owner)
	f.SetCellValue(infoSheet, "B2", u.repo)
	f.SetCellValue(infoSheet, "C2", from)
	f.SetCellValue(infoSheet, "D2", to)
	f.SetCellValue(infoSheet, "E2", u.develBranch)
	f.SetCellValue(infoSheet, "F2", u.masterBranch)

	tableString := &strings.Builder{}
	table := tablewriter.NewWriter(tableString)
	table.SetHeader([]string{"PR", "Commits", "Size", "Time To First Review", "Review time", "Last Review To Merge", "Comments", "PR Lead Time", "Time To Merge"})

	sheetName := "Feature Branch"
	f.NewSheet(sheetName)
	f.SetColWidth(sheetName, "A", "J", 18)
	f.SetCellValue(sheetName, "A1", "PR")
	f.SetCellValue(sheetName, "B1", "Commits")
	f.SetCellValue(sheetName, "C1", "Size")
	f.SetCellValue(sheetName, "D1", "Time To First Review")
	f.SetCellValue(sheetName, "E1", "Review time")
	f.SetCellValue(sheetName, "F1", "Last Review To Merge")
	f.SetCellValue(sheetName, "G1", "Comments")
	f.SetCellValue(sheetName, "H1", "PR Lead Time")
	f.SetCellValue(sheetName, "I1", "Time To Merge")
	f.SetCellValue(sheetName, "J1", "Date")

	row := 2

	for _, pr := range prs {
		str1, _ := durationfmt.Format(pr.TimeToFirstReview(), "%dd %hh %mm")
		str2, _ := durationfmt.Format(pr.TimeToReview(), "%dd %hh %mm")
		str3, _ := durationfmt.Format(pr.LastReviewToMerge(), "%dd %hh %mm")
		str4, _ := durationfmt.Format(pr.PRLeadTime(), "%dd %hh %mm")
		str5, _ := durationfmt.Format(pr.TimeToMerge(), "%dd %hh %mm")

		f.SetCellValue(sheetName, fmt.Sprintf("%s%d", "A", row), pr.Number)
		f.SetCellValue(sheetName, fmt.Sprintf("%s%d", "B", row), pr.Commits)
		f.SetCellValue(sheetName, fmt.Sprintf("%s%d", "C", row), pr.ChangedLines)
		f.SetCellValue(sheetName, fmt.Sprintf("%s%d", "D", row), str1)
		f.SetCellValue(sheetName, fmt.Sprintf("%s%d", "E", row), str2)
		f.SetCellValue(sheetName, fmt.Sprintf("%s%d", "F", row), str3)
		f.SetCellValue(sheetName, fmt.Sprintf("%s%d", "G", row), pr.ReviewComments)
		f.SetCellValue(sheetName, fmt.Sprintf("%s%d", "H", row), str4)
		f.SetCellValue(sheetName, fmt.Sprintf("%s%d", "I", row), str5)
		f.SetCellValue(sheetName, fmt.Sprintf("%s%d", "J", row), pr.FirstCommitAt)

		table.Append([]string{

			strconv.Itoa(pr.Number),
			strconv.Itoa(pr.Commits),
			strconv.Itoa(pr.ChangedLines),
			str1, str2, str3,
			//DurationFormater(pr.TimeToFirstReview()),
			//DurationFormater(pr.TimeToReview()),
			//DurationFormater(pr.LastReviewToMerge()),
			strconv.Itoa(pr.ReviewComments),
			str4, str5,
			//DurationFormater(pr.PRLeadTime()),
			//DurationFormater(pr.TimeToMerge()),
		})

		row++
	}

	kpi := vcs.NewKPICalculator(prs)

	table.SetFooter([]string{
		fmt.Sprintf("Count: %d", kpi.CountPR()),
		fmt.Sprintf("AVG: %.2f", kpi.AvgCommits()),
		fmt.Sprintf("AVG: %.2f", kpi.AvgChangedLines()),
		AvgDurationFormater(kpi.AvgTimeToFirstReview()),
		AvgDurationFormater(kpi.AvgTimeToReview()),
		AvgDurationFormater(kpi.AvgLastReviewToMerge()),
		fmt.Sprintf("AVG: %.2f", kpi.AvgReviews()),
		AvgDurationFormater(kpi.AvgPRLeadTime()),
		AvgDurationFormater(kpi.AvgTimeToMerge()),
	}) // Add Footer

	table.SetAlignment(tablewriter.ALIGN_LEFT)
	table.SetBorder(false)
	table.Render() // Send output
	return tableString.String(), nil
}

func (u CmdUI) getReleaseBranchReport(from, to time.Time) (string, error) {
	prs, err := u.client.GetMergedPRList(u.owner, u.repo, from, to, u.masterBranch)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error gathering information: %s", err.Error())
		return "", err
	}

	tableString := &strings.Builder{}
	table := tablewriter.NewWriter(tableString)
	table.SetHeader([]string{"PR", "Commits", "Size", "PR Lead Time", "Time To Merge"})

	sheetName2 := "Master Branch"
	f.NewSheet(sheetName2)
	f.SetColWidth(sheetName2, "A", "F", 18)
	f.SetCellValue(sheetName2, "A1", "PR")
	f.SetCellValue(sheetName2, "B1", "Commits")
	f.SetCellValue(sheetName2, "C1", "Size")
	f.SetCellValue(sheetName2, "D1", "PR Lead Time")
	f.SetCellValue(sheetName2, "E1", "Time To Merge")
	f.SetCellValue(sheetName2, "F1", "Date")

	row2 := 2

	for _, pr := range prs {
		str1, _ := durationfmt.Format(pr.PRLeadTime(), "%dd %hh %mm")
		str2, _ := durationfmt.Format(pr.TimeToMerge(), "%dd %hh %mm")

		f.SetCellValue(sheetName2, fmt.Sprintf("%s%d", "A", row2), pr.Number)
		f.SetCellValue(sheetName2, fmt.Sprintf("%s%d", "B", row2), pr.Commits)
		f.SetCellValue(sheetName2, fmt.Sprintf("%s%d", "C", row2), pr.ChangedLines)
		f.SetCellValue(sheetName2, fmt.Sprintf("%s%d", "D", row2), str1)
		f.SetCellValue(sheetName2, fmt.Sprintf("%s%d", "E", row2), str2)
		f.SetCellValue(sheetName2, fmt.Sprintf("%s%d", "F", row2), pr.FirstCommitAt)

		table.Append([]string{
			strconv.Itoa(pr.Number),
			strconv.Itoa(pr.Commits),
			strconv.Itoa(pr.ChangedLines),
			str1, str2,
			//DurationFormater(pr.PRLeadTime()),
			//DurationFormater(pr.TimeToMerge()),
		})

		row2++
	}

	kpi := vcs.NewKPICalculator(prs)

	table.SetFooter([]string{
		fmt.Sprintf("Count: %d", kpi.CountPR()),
		fmt.Sprintf("AVG: %.2f", kpi.AvgCommits()),
		fmt.Sprintf("AVG: %.2f", kpi.AvgChangedLines()),
		AvgDurationFormater(kpi.AvgPRLeadTime()),
		AvgDurationFormater(kpi.AvgTimeToMerge()),
	}) // Add Footer

	table.SetAlignment(tablewriter.ALIGN_LEFT)
	table.SetBorder(false)
	table.Render() // Send output
	return tableString.String(), nil
}
