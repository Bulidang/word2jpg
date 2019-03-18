$PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition

function word2pdf {
    param ($documents_path)

    if (-not (test-path "$PSScriptRoot\jpg")) {
        mkdir "$PSScriptRoot\jpg"
    }
    else {
        Remove-Item "$PSScriptRoot\jpg\*" -Recurse
    }

    $word_app = New-Object -ComObject Word.Application
    Get-ChildItem -Path $documents_path -Filter *.doc? -r | ForEach-Object {
        $document = $word_app.Documents.Open($_.FullName)
        $pdf_fullname = "$($PSScriptRoot)\jpg\$($_.BaseName).pdf"
        $document.SaveAs($pdf_fullname, 17)
        $document.Close()
        write-host $pdf_fullname
    }
    $word_app.Quit()
}


function pdf2jpg {
    param ($documents_path)

    Get-ChildItem $documents_path *.pdf -r  | ForEach-Object {   
        $pdf_fullname = "$($_.DirectoryName)\$($_.basename).pdf"
        $jpg_path = "$PSScriptRoot\jpg\$($_.basename).jpg"
        write-host $jpg_path
        $exe_path = "$PSScriptRoot\bin\gswin32c.exe"
        & $exe_path -dNOPAUSE -sDEVICE=jpeg -o "$jpg_path" -dFirstPage=1  -dJPEGQ=80 -r300  -f "$pdf_fullname"  -c quit
        write-host $jpg_path
        
        Remove-Item $pdf_fullname
    }
}


if ($args.count -eq 1) {
    write-host "Convert doc(x) files in  '$args' directory to jpg:" -ForegroundColor Green -BackgroundColor Black
    word2pdf($args)
    pdf2jpg("$PSScriptRoot\jpg")
}
else {
    write-host "Need only one Parameter: -Path documents_path" -ForegroundColor Red -BackgroundColor Black
}





