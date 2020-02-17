# Powershell script to generate APA formatted Word Document
# Written by Joshua Woleben
# 9/15/2019

# Variable declarations

$margins = 72 * 1.0 # in inches
$running_header_title = "RUNNING HEADER: "
$paper_title = ""
$paper_abstract_title = "Abstract"
$paper_author = ""
$references_title = "References"
$institution_name = ""
$font = "Times New Roman"
$title_page_font_size = "16"
$author_lines_font_size = "14"
$body_font_size = "12"
$start_page_number = 2


# GUI Code
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="APA Formatter" Height="1100" Width="450" MinHeight="500" MinWidth="400" ResizeMode="CanResizeWithGrip">
    <StackPanel>
        <Label x:Name="Title" Content="Paper Title"/>
        <TextBox x:Name="TitleTextBox"/>
        <Label x:Name="Author" Content="Paper Author (First M Last)"/>
        <TextBox x:Name="AuthorTextBox"/>
        <Label x:Name="InstructorName" Content="Instructor Name"/>
        <TextBox x:Name="InstructorTextBox"/>
        <Label x:Name="CourseName" Content="Course Name"/>
        <TextBox x:Name="CourseTextBox"/>
        <Label x:Name="Institution" Content="Institution Name"/>
        <TextBox x:Name="InstitutionTextBox"/>
        <Label x:Name="AbstractText" Content="Abstract Text"/>
        <TextBox x:Name="AbstractTextBox" Height="100" AcceptsReturn="true" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"/>
        <Label x:Name="BodyText" Content="Main Body Text"/>
        <TextBox x:Name="MainBodyTextBox" Height="300" AcceptsReturn="true" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"/>
        <Label x:Name="ReferencesText" Content="References Text"/>
        <TextBox x:Name="ReferencesTextBox" Height="200" AcceptsReturn="true" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"/>
        <Label x:Name="FilePath" Content="Word Doc Save Path"/>
        <TextBox x:Name="WordPathTextBox"/>
        <Button x:Name="FilePickerButton" Content="Set File Path" Margin="10,10,10,0" VerticalAlignment="Top" Height="25"/>
        <Button x:Name="FormatButton" Content="Format APA!" Margin="10,10,10,0" VerticalAlignment="Top" Height="25"/>
    </StackPanel>
</Window>
'@
 
$global:Form = ""
# XAML Launcher
$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
try{$global:Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Some possible causes for this problem include: .NET Framework is missing PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."; break}
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name ($_.Name) -Value $global:Form.FindName($_.Name)}

# Set up controls
$TitleTextBox = $global:Form.FindName('TitleTextBox')
$AuthorTextBox = $global:Form.FindName('AuthorTextBox')
$InstitutionTextBox = $global:Form.FindName('InstitutionTextBox')
$CourseTextBox = $global:Form.FindName('CourseTextBox')
$InstructorTextBox = $global:Form.FindName('InstructorTextBox')
$AbstractTextBox = $global:Form.FindName('AbstractTextBox')
$MainBodyTextBox = $global:Form.FindName('MainBodyTextBox')
$ReferencesTextBox = $global:Form.FindName('ReferencesTextBox')
$WordPathTextBox = $global:Form.FindName('WordPathTextBox')
$FilePickerButton = $global:Form.FindName('FilePickerButton')
$FormatButton = $global:Form.FindName('FormatButton')

$FilePickerButton.Add_Click({
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.SaveFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('CommonDocuments') }
    $null = $FileBrowser.ShowDialog()
    $WordPathTextBox.Text = $FileBrowser.FileName
})
$FormatButton.Add_Click({


Add-Type -AssemblyName "Microsoft.Office.Interop.Word"

# Get user parameters
$paper_title = $TitleTextBox.Text
$paper_author = $AuthorTextBox.Text
$institution_name = $InstitutionTextBox.Text
$word_path = $WordPathTextBox.Text
$course_name = $CourseTextBox.Text
$instructor_name = $InstructorTextBox.Text

# Error checking
if ([string]::IsNullOrWhiteSpace($paper_title)) {
    [System.Windows.MessageBox]::Show('No title specified!')
    return
}
if ([string]::IsNullOrWhiteSpace($paper_author)) {
    [System.Windows.MessageBox]::Show('No author specified!')
    return
}
if ([string]::IsNullOrWhiteSpace($paper_title)) {
    [System.Windows.MessageBox]::Show('No institution specified!')
    return
}
if ([string]::IsNullOrWhiteSpace($word_path)) {
    [System.Windows.MessageBox]::Show('No Word save path specified!')
    return
}

$index = 0

# Open Word object
$word_object = New-Object -ComObject Word.Application
$word_object.Visible = $true

# Add word document
$word_document = $word_object.Documents.Add()
$current_selection = $word_object.Selection

# Set up basic formatting
$current_selection.Font.Name = $font
$word_document.PageSetup.LeftMargin = $margins
$word_document.PageSetup.RightMargin = $margins
$word_document.PageSetup.TopMargin = $margins
$word_document.PageSetup.BottomMargin = $margins

# Create header
$word_document.PageSetup.DifferentFirstPageHeaderFooter = $true
$header = $word_document.Sections(1).Headers([Microsoft.Office.Interop.Word.WdHeaderFooterIndex]::wdHeaderFooterFirstPage)
$header.Range.Text = ($running_header_title + $paper_title)

# Create title page
$current_selection.Font.Size = $title_page_font_size
$current_selection.TypeText("`v`v`v")
$current_selection.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.wdParagraphAlignment]::wdAlignParagraphCenter # Set alignment to center
$current_selection.TypeText($paper_title)
$current_selection.TypeText("`v")
$current_selection.Font.Size = $author_lines_font_size
$current_selection.TypeText($paper_author)
$current_selection.TypeText("`v")
$current_selection.Typetext($instructor_name)
$current_selection.Typetext("`v")
$current_selection.Typetext($course_name)
$current_selection.Typetext("`v")
$current_selection.Typetext($institution_name)
$current_selection.Typetext("`v")
$current_selection.Typetext($(get-date -Format 'MMMM d, yyyy'))
$current_selection.Typetext("`v")


# Insert page break
$current_selection.InsertNewPage()


# Create abstract page
$current_selection.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.wdParagraphAlignment]::wdAlignParagraphCenter # Set alignment to center
$current_selection.ParagraphFormat.LineSpacingRule = 2 # Set double spacing
$current_selection.Font.Size = $body_font_size

# Only insert abstract if one is present, otherwise skip
if (-not [string]::IsNullOrWhiteSpace($AbstractTextBox.Text)) {
    $current_selection.TypeText($paper_abstract_title)
    $current_selection.TypeParagraph()
    $index++

    # Get abstract text
    $current_selection.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.wdParagraphAlignment]::wdAlignParagraphLeft # Set alignment to left
    $abstract_text = ($AbstractTextBox.Text).Trim()
    $current_selection.TypeText($abstract_text)
    $index++

    # Insert page break
    $current_selection.InsertNewPage()
}

# Process main body section
$current_selection.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.wdParagraphAlignment]::wdAlignParagraphLeft # Set alignment to left
$current_selection.ParagraphFormat.CharacterUnitFirstLineIndent = 5
# Get main body text
$main_body_text = ($MainBodyTextBox.Text).Trim()

# Process main body text
foreach ($line in $main_body_text) {
    $current_selection.TypeText($line)
    if ($line -match "\n") {
        $current_selection.TypeParagraph()
        $index++
    }
}

# Insert page break
$current_selection.InsertNewPage()

# Create References page
$current_selection.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.wdParagraphAlignment]::wdAlignParagraphCenter # Set alignment to center
$current_selection.TypeText($references_title)
$current_selection.TypeParagraph()

$current_selection.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.wdParagraphAlignment]::wdAlignParagraphLeft # Set alignment to left

# Get references text
$references_text = ($ReferencesTextBox.Text).Trim()

# Set hanging indent for references
# $current_selection.Paragraph.FirstLineIndent = -1

#$current_selection.InsertBreak([Microsoft.Office.Interop.Word.wdBreakType]::wdSectionBreakContinuous)
$current_selection.ParagraphFormat.CharacterUnitFirstLineIndent = -5
# Process references
foreach ($line in $references_text) {
   $current_selection.TypeText($line)
    if ($line -match "\n") {
        $current_selection.TypeParagraph()
    }
}


# Rest of header pages
$new_header = $word_document.Sections(1).Headers([Microsoft.Office.Interop.Word.WdHeaderFooterIndex]::wdHeaderFooterPrimary)
$new_header.Range.Text = $paper_title

# add page numbers
$word_document.Sections(1).Headers([Microsoft.Office.Interop.Word.WdHeaderFooterIndex]::wdHeaderFooterPrimary).PageNumbers.Add(2) # Add page numbers to right header

$word_document.SaveAs($word_path,[ref][Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault)

$word_object.Quit()

$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word_object)
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
Remove-Variable word_object
[System.Windows.MessageBox]::Show('Completed and saved to ' + $word_path + '!')
}) # End click button

# Show GUI
$global:Form.ShowDialog() | out-null