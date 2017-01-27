#Requires -Version 5.0

<#

.SYNOPSIS

    Turning object into the HTML table using downloaded datatables.net


.DESCRIPTION

    This function convert your object to HTML using JS and DataTables.net.
    
    Features:
        * sort columns
        * dynamically global search
        * dynamically filter per columns
        * set row color depends on 'status' row (critical|stopped=red; OK|running=green; warning=yellow)
        * export to pdf/excel/csv and copy/print
        * save state (like order settings)

    How to start?
        1) Download datatables (https://datatables.net/download/packages), unpack it and save to for example c:\WWW\datatables\ (there should be folders like 'examples', extensions' and 'media')
        2) Launch any web server you have. For Example Mongoose Web Server v6.5 (https://www.cesanta.com/products/binary)
        3) When Executing your script call this function first:
         . "C:\scripts\Out-HTMLDataTable.ps1"
        4) $Path should save file to c:\WWW\table01.htm
        5) Open in webbrowser: http://127.0.0.1:8080/table01.htm

    How folder structire should looks like?
        .\www\
              datatables\examples
              datatables\extensions
              datatables\media
              icons\favicon.png
        .\table01.htm


.PARAMETER InputObject
    Just object (but not via pipeline so far). Example value:
    (get-service | select Name, StartType, Status)


.PARAMETER Path
    Path where will be html report saved. Example value:
    'c:\WWW\table01.htm'


.PARAMETER Title
    Insert your main title for report. Example value:
    'Services report'


.PARAMETER RefreshEvery
    This is just info on html report that this report is refreshing every X min/hours/days.
    This is only text. It depends on your task scheduler configuration. Example value:
    '15 min.'

.PARAMETER BodyContent
    Write down your description for this report, table, rows or from where data is come from. Example value:
    'Below are services on machine msfile01'


.PARAMETER Author
    Info about author. Just write your name here. Example value:
    'Me and my hamster'


.PARAMETER Culture
    Sometimes it's needed to set culture for datetime format. You can do it here. Example value:
    'pl-pl'


.PARAMETER stopwatch
    You can measure time needed to generate data by your script to until this function gets it. Example value:
    In your script on beggining:
        $Stopwatch	= [System.Diagnostics.Stopwatch]::StartNew()
    And when you call this functions:
        $stopwatch


.PARAMETER StatusRow
    If you have a row/column with statuses like warning/critical/OK or Running/Stopped you can enable this parameter.
    Remember that row start from 0.
    So that if you have: (get-service | select Name, StartType, Status) and StatusRow 2 means the StatusRow will be column named "Status".
    Example value:
    2


.PARAMETER OrderRow
    You can set column number for ordering. Remember that row start from 0.
    Read more at: https://datatables.net/examples/basic_init/table_sorting.html
    Example value:
    2
    (You can combine it with parameter OrderDirection.)


.PARAMETER OrderDirection
    If you used OrderRow then you can use this parameter. asc (default) od desc.
    Read more at: https://datatables.net/examples/basic_init/table_sorting.html
    Example value:
    desc


.PARAMETER StateSave
    Enable or disable state saving. Read more at: https://datatables.net/reference/option/stateSave . Example value:
    $true 


.PARAMETER ExportButtons
    Enable or disable export buttons. Read more at: https://datatables.net/extensions/buttons/examples/initialisation/export.html
    Example value:
    $true


.PARAMETER RowSelection
    Enable or disable row selection (on click). read more: https://datatables.net/examples/server_side/select_rows.html
    Example value:
    $true


.PARAMETER IconPath
    You can define image path for site (in browser bar). PNG Format. Example value:
    'icons/favicon.png'


.EXAMPLE #1

    The most simplest example:
    Out-HTMLDataTable -InputObject (get-service | select Name, StartType, Status) -Path 'c:\WWW\table01.htm'


.EXAMPLE #2

    More advanced:
    $Stopwatch	= [System.Diagnostics.Stopwatch]::StartNew()
    Out-HTMLDataTable -InputObject (get-service | select Name, StartType, Status) -Path 'c:\WWW\table01.htm' `
        -Title 'Services report' -RefreshEvery '15 minutes' -BodyContent 'Below are services on machine msfile01' `
        -Author 'Me and my hamster' -Culture 'pl-pl' -stopwatch $stopwatch -StatusRow 2 `
        -OrderRow 2 -OrderDirection 'desc' -StateSave $false -ExportButtons $true -RowSelection $true `
        -IconPath 'icons/favicon.png'

.NOTES

    Version:        1.0
    Author:         Lukasz Wasko
    Creation Date:  2017.01.25

    Notice, that some files are downloaded from Internet:
    * https://code.jquery.com/jquery-1.12.4.js
    * https://cdnjs.cloudflare.com/ajax/libs/jszip/2.5.0/jszip.min.js
    * https://cdn.rawgit.com/bpampuch/pdfmake/0.1.18/build/pdfmake.min.js
    * https://cdn.rawgit.com/bpampuch/pdfmake/0.1.18/build/vfs_fonts.js
    So if you want to open reports on systems, that do not have Internet access
    just download these files and change Intenret paths to local paths in function.

    Tested on:
    * Windows 10
    * Powershell 5.1
    * DataTables v1.10.13
    * free Mongoose Web Server v6.5

    Sorry for all that "`n". I just want to keep clear formatting in htm file

#>

function Out-HTMLDataTable {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true,
                   Position = 1)]
            [PSObject]
            $InputObject,
        [Parameter(Mandatory = $true,
                   ValueFromPipeline = $false,
                   Position = 2)]
            [string]
            $Path,
        [Parameter(Mandatory = $false,
                   ValueFromPipeline = $false)]
            [string]
            $Title,
        [Parameter(Mandatory = $false)]
            [string]
            $BodyContent,
        [Parameter(Mandatory = $false)]
            [string]
            $RefreshEvery,
        [Parameter(Mandatory = $false,
                   ValueFromPipeline = $false)]
            [string]
            $Culture,
        [Parameter(Mandatory = $false,
                   ValueFromPipeline = $false)]
            [Diagnostics.Stopwatch]
            $Stopwatch,
        [Parameter(Mandatory = $false,
                   ValueFromPipeline = $false)]
            [string]
            $Author,
        [Parameter(Mandatory = $false,
                   ValueFromPipeline = $false)]
            [byte]
            $StatusRow,
        [Parameter(Mandatory = $false,
                   ValueFromPipeline = $false)]
            [byte]
            $OrderRow,
        [validateset('asc', 'desc')]
            [string]
            $OrderDirection = 'asc',
        [Parameter(Mandatory = $false,
                   ValueFromPipeline = $false)]
            [boolean]
            $StateSave = $false,
        [Parameter(Mandatory = $false,
                   ValueFromPipeline = $false)]
            [boolean]
            $RowSelection = $false,
        [Parameter(Mandatory = $false,
                   ValueFromPipeline = $false)]
            [boolean]
            $ExportButtons = $false,
        [Parameter(Mandatory = $false,
                   ValueFromPipeline = $false)]
            [string]
            $IconPath
        
    )

    BEGIN {
        
        if( !(Test-Path "$( (get-item $Path).DirectoryName )\Datatables") )
        {
            Write-Warning -Message "Datatables not found in: $( (get-item $Path).DirectoryName )\Datatables folder"
            Write-Warning -Message "Download it from: https://datatables.net"
        }
        else
        {
            Write-Verbose -Message "DataTables found."
        }

        Write-Verbose -Message 'Generating column names..' #W późniejszym kroku bedą służyły jako indywidualne pola Search dla każdej kolumny
        $TableFoot = Foreach ($prop in ('NoteProperty', 'Property', 'ScriptProperty')) {
            if($InputObject | Get-Member -MemberType $prop)
            {
                ($InputObject | Get-Member -MemberType $prop) | Foreach {'<th>' + $PSItem.name + '</th>'} 
                break
            }
        }
        Write-Verbose -Message " Selected property: $prop"
        Write-Debug -Message " Found: $($TableFoot.count) columns."
        
        Write-Verbose -Message 'Preparing html HEAD..'
        $head = '    <meta charset="utf-8">' + "`n" +
        $(
            if($Title)
            {
                "    <title>$Title</title>" + "`n"
                Write-Verbose -Message ' Title: Added.'
            }
            if($IconPath)
            {
                '   <link rel="shortcut icon" type="image/png" href="' + $IconPath + '">' + "`n"
                Write-Verbose -Message ' IconPath: Added.'
                
                if( !(Test-Path "$( (get-item $Path).DirectoryName )\$IconPath") )
                {
                    Write-Warning -Message "  Icon file not found in: $( (get-item $Path).DirectoryName )\$IconPath"
                }
                else
                {
                    Write-Verbose -Message "  Icon file found."
                }
            }
        ) +
        '    <meta name="viewport" content="initial-scale=1.0, maximum-scale=2.0">' + "`n" +
        '    <link rel="stylesheet" type="text/css" href="datatables/media/css/jquery.dataTables.css">' + "`n" +
	    '    <link rel="stylesheet" type="text/css" href="datatables/examples/resources/syntax/shCore.css">' + "`n" +
	    '    <link rel="stylesheet" type="text/css" href="datatables/examples/resources/demo.css">' + "`n" +
        $(
            if($ExportButtons)
            {
                '    <link rel="stylesheet" type="text/css" href="datatables/extensions/Buttons/css/buttons.dataTables.min.css">' + "`n"
                Write-Verbose -Message ' ExportButtons css: Added.'
            }
        ) +
	    '    <style type="text/css" class="init">' + "`n" +
        '    </style>' + "`n" +
	    '    <script type="text/javascript" language="javascript" src="//code.jquery.com/jquery-1.12.4.js">' + "`n" +
	    '    </script>' + "`n" +
	    '    <script type="text/javascript" language="javascript" src="datatables/media/js/jquery.dataTables.js">' + "`n" +
	    '    </script>' + "`n" +
	    '    <script type="text/javascript" language="javascript" src="datatables/examples/resources/syntax/shCore.js">' + "`n" +
	    '    </script>' + "`n" +
        '    <script type="text/javascript" language="javascript" src="datatables/examples/resources/demo.js">' + "`n" +
	    '    </script>' + "`n" +
        $( 
            if($ExportButtons)
            {
                '    <script type="text/javascript" language="javascript" src="datatables/extensions/Buttons/js/dataTables.buttons.min.js">' + "`n" +
	            '    </script>' + "`n" +
                '    <script type="text/javascript" language="javascript" src="datatables/extensions/Buttons/js/buttons.flash.min.js">' + "`n" +
	            '    </script>' + "`n" +
                '    <script type="text/javascript" language="javascript" src="datatables/extensions/Buttons/js/buttons.html5.min.js">' + "`n" +
	            '    </script>' + "`n" +
                '    <script type="text/javascript" language="javascript" src="datatables/extensions/Buttons/js/buttons.print.min.js">' + "`n" +
	            '    </script>' + "`n" +
                '    <script type="text/javascript" language="javascript" src="https://cdnjs.cloudflare.com/ajax/libs/jszip/2.5.0/jszip.min.js">' + "`n" +
	            '    </script>' + "`n" +
                '    <script type="text/javascript" language="javascript" src="https://cdn.rawgit.com/bpampuch/pdfmake/0.1.18/build/pdfmake.min.js">' + "`n" +
	            '    </script>' + "`n" +
                '    <script type="text/javascript" language="javascript" src="https://cdn.rawgit.com/bpampuch/pdfmake/0.1.18/build/vfs_fonts.js">' + "`n" +
	            '    </script>' + "`n"
                Write-Verbose -Message ' ExportButtons js files: Added.'
            }
        ) +
	    '    <script type="text/javascript" language="javascript" class="init">' + "`n" +
        
        '        $(document).ready(function() {' + "`n" +
        '                        // Setup - add a text input to each footer cell' + "`n" +
        '                        ' + "`n" +
        '                        $("#table01 tfoot th").each( function () {' + "`n" +
        '                        	var title = $(this).text();' + "`n" +
        "                        	`$(this).html('<input type=`"text`" placeholder=`"Search '+title+'`" />' );" + "`n" +
        '                        } );' + "`n" +
        '                        // DataTable' + "`n" +
        "                        var table = `$('#table01').DataTable(" + "`n" +
        '                        {' + "`n" +
        '                        	"paging":   false,' + "`n" +
        $(
            if($ExportButtons)
            {
                "                            dom: 'Bfrtip'," + "`n"
                '                           buttons: [' + "`n"
                "                            	'copy', 'csv', 'excel', 'pdf', 'print'" + "`n"
                "                           ]," + "`n"
                Write-Verbose -Message ' ExportButtons DT options: Added.'
            }
            if($orderRow)
            {
                "                           `"order`": [[ $OrderRow, `"$OrderDirection`" ]]," + "`n"
                Write-Verbose -Message ' OrderRow DT options: Added.'
            }
            if($StateSave)
            {
                "                       `"stateSave`": true," + "`n"
                Write-Verbose -Message ' StateSave DT options: Added.'
            }
            if($StatusRow)
            {
                '                           "fnRowCallback": function( nRow, aData, iDisplayIndex, iDisplayIndexFull ) {' + "`n" +
                "                                if ( aData[$StatusRow].toLowerCase() == `"critical`" || aData[$StatusRow].toLowerCase() == `"stopped`" )" + "`n" +
                '                                {' + "`n" +
                '                                $("td", nRow).css("background-color", "#FFCCCC");' + "`n" +
                '                                }' + "`n" +
                "                                else if ( aData[$StatusRow].toLowerCase() == `"warning`")" + "`n" +
                '                                {' + "`n" +
                '                                   $("td", nRow).css("background-color", "#FFFF99");' + "`n" +
                '                                }' + "`n" +
                "                                else if ( aData[$StatusRow].toLowerCase() == `"ok`" || aData[$StatusRow].toLowerCase() == `"running`" )" + "`n" +
                '                                {' + "`n" +
                '                                   $("td", nRow).css("background-color", "#CCFFCC");' + "`n" +
                '                                }' + "`n" +
                '                            }'
                Write-Verbose -Message ' StatusRow DT options: Added.'
            }
        ) + "`n" +
        '                        } );' + "`n" +
        '				// Apply the search' + "`n" +
        '				table.columns().every( function () {' + "`n" +
        '					var that = this;' +"`n" +
        "					`$( 'input', this.footer() ).on( 'keyup change', function () {" + "`n" +
        '						if ( that.search() !== this.value ) {' +"`n" +
        '							that' +"`n" +
        '								.search( this.value )' +"`n" +
        '								.draw();' +"`n" +
        '						}' +"`n" +
        '					} );' +"`n" +
        '				} );' + "`n" +
        $(
            if($RowSelection)
            {
                "				`$('#table01 tbody').on( 'click', 'tr', function () {" + "`n" +
				"					`$(this).toggleClass('selected');" + "`n" +
				"				} );" + "`n"
                Write-Verbose -Message ' RowSelection DT options: Added.'
            }
        ) +
        '        } );' +"`n" +
        '    </script>' + "`n"

        
        Write-Verbose -Message 'Preparing html BODY..'
        $body = '<font face="verdana" size="2">' + "`n" +
            $(
                if ($Title)
                {
                    "<H1>$Title</H1>" + "`n"
                    Write-Verbose -Message ' Title: Added.'
                }
            ) +
            '<small>Generated at:&#9; ' +
            $(
                if ($Culture)
                {
                    (get-date).ToString('G',$(New-Object globalization.cultureinfo($Culture)))
                    Write-Verbose -Message " Culture: Added. [$Culture]"
                }
                else
                {
                    (get-date).ToString('G')
                    " Culture: Added. [$((Get-Culture).name)]"
                }
            ) + '&#9;&#9;' + "`n" +
            $(
                if ($RefreshEvery)
                {
                    "(Refresh every: $RefreshEvery)<br>" + "`n"
                    Write-Verbose -Message ' RefreshEvery: Added.'
                }
                if ($BodyContent)
                {
                    $BodyContent + "`n"
                    Write-Verbose -Message ' BodyContent: Added.'
                }
            ) + "`n" + '</small>'


        Write-Verbose -Message 'Preparing html POSTCONTENT..'
        $PostContent = '<br><HR>' + "`n" +
            $(
                if($Stopwatch)
                {
                    "<font size='2'>The task took: $([system.String]::Format('{0:00}h {1:00}m {2:00}s', $Stopwatch.Elapsed.Hours, $Stopwatch.Elapsed.Minutes, $Stopwatch.Elapsed.Seconds);)<br></font>"
                    Write-Verbose -Message ' StopWatch: Added.'
                }
            ) +
            $(
                if($Author)
                {
                    "<font size='2'>Author: $Author</font>"
                    Write-Verbose -Message ' Author: Added.'
                }
            )
        
        Write-Verbose -Message 'Converting to HTML and saving to file..'
        Write-Debug " Output Path: $Path"
        $InputObject | ConvertTo-Html -head $Head -body $Body -PostContent $PostContent | Out-string |
	        ForEach-Object {
                $PSItem.replace('&lt;','<').replace('&gt;','>').replace('<tr><th>',"<table id='table01' class='display compact' cellspacing='0' width='100%'><thead><tr><th>").replace('</th></tr>','</th></tr></thead><tbody>').replace('</table>','</tbody></table>').replace('</thead>', $('</thead><tfoot><tr>' + $TableFoot + '</tr></tfoot>'))
            } | Out-File $Path -Encoding UTF8 -Force
    }
}
