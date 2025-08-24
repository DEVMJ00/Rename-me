Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Création de la fenêtre
$form = New-Object System.Windows.Forms.Form
$form.Text = "Rename-me"
$form.Size = New-Object System.Drawing.Size(492, 300)
# Définir la police par défaut pour tous les contrôles
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255, 255)
$form.StartPosition = "CenterScreen"

# Texte d'instruction : Selection du mail
$labelSelectMail = New-Object System.Windows.Forms.Label
$labelSelectMail.Text = "Sélectionner le fichier à renommer :"
$labelSelectMail.AutoSize = $true
$labelSelectMail.Location = New-Object System.Drawing.Point(20, 20)
$form.Controls.Add($labelSelectMail)

# Liste des fichiers .msg
$comboFiles = New-Object System.Windows.Forms.ComboBox
$comboFiles.Location = New-Object System.Drawing.Point(20, 50)
$comboFiles.Size = New-Object System.Drawing.Size(300, 20)
$comboFiles.DropDownStyle = 'DropDownList'
$form.Controls.Add($comboFiles)


# Texte d'instruction : Selection du nouveau nom
$labelSelectNewName = New-Object System.Windows.Forms.Label
$labelSelectNewName.Text = "Choisir un nouveau nom :"
$labelSelectNewName.AutoSize = $true
$labelSelectNewName.Location = New-Object System.Drawing.Point(20, 95)
$form.Controls.Add($labelSelectNewName)

# Liste des options de renommage
$comboOptions = New-Object System.Windows.Forms.ComboBox
$comboOptions.Location = New-Object System.Drawing.Point(20, 120)
$comboOptions.Size = New-Object System.Drawing.Size(300, 20)
$comboOptions.DropDownStyle = 'DropDownList'
$comboOptions.DropDownHeight = 300  # hauteur du menu déroulant en pixels
$comboOptions.Items.AddRange(@(
    "1 - 1ère relance",
    "2 - 2ème relance",
    "3 - 3ème relance",
    "4 - 4ème relance",
    "5 - Information complémentaire",
    "6 - Demande de devis",
    "7 - Devis",
    "8 - Demande d'intervention",
    "9 - Rapport d'intervention",
    "0 - Facture",
    
    "R - Réponse"
))
$form.Controls.Add($comboOptions)

# Bouton de renommage
$btnRenommer = New-Object System.Windows.Forms.Button
$btnRenommer.Text = "Renommer"
$btnRenommer.Location = New-Object System.Drawing.Point(350, 120)
$btnRenommer.Size = New-Object System.Drawing.Size(100, 25)
$form.Controls.Add($btnRenommer)


# Label d'information (zone de texte en bas)
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.AutoSize = $true
$labelStatus.Location = New-Object System.Drawing.Point(20, 160)
$labelStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.Add($labelStatus)

# Label Footer
$labelFooter = New-Object System.Windows.Forms.Label
$labelFooter.Text = "Made with ❤ by DEVMJ"
#$labelFooter.ForeColor = [System.Drawing.Color]::Gainsboro
$labelFooter.ForeColor = [System.Drawing.Color]::FromArgb(50, 179, 179, 179)  
$labelFooter.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Regular)
$labelFooter.AutoSize = $true
#$labelFooter.Location = New-Object System.Drawing.Point(300,492)
# Centrage horizontal et vertical
$LabelFooter.Location = New-Object System.Drawing.Point(
    (($Form.ClientSize.Width - $LabelFooter.Width) / 2),
    ($Form.ClientSize.Height - $LabelFooter.Height - 1)
)
# Création du tooltip
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($labelFooter, "En savoir plus")

# Curseur en forme de main
$LabelFooter.Cursor = [System.Windows.Forms.Cursors]::Hand

# Action au clic -> ouvre GitHub
$LabelFooter.Add_Click({
    Start-Process "https://github.com/DEVMJ00/Rename-me"
})

$form.Controls.Add($labelFooter)


# Création du tooltip
$toolTip = New-Object System.Windows.Forms.ToolTip

# Chargement des fichiers .msg à l'ouverture
$script:msgFiles = Get-ChildItem -Path "." -Filter "*.msg" | Select-Object -ExpandProperty Name
if ($msgFiles.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show("Aucun fichier .msg existant dans le dossier.","Information",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
    $form.Close()
}
else {
    # Remplir le ComboBox
    $comboFiles.Items.AddRange($msgFiles)
    $comboFiles.SelectedIndex = 0
    $comboOptions.SelectedIndex = 0


    # Initialiser le tooltip avec l’élément sélectionné
    $toolTip.SetToolTip($comboFiles, $comboFiles.SelectedItem)

    # Mettre à jour le tooltip quand la sélection change
    $comboFiles.Add_SelectedIndexChanged({
        $toolTip.SetToolTip($comboFiles, $comboFiles.SelectedItem)
    })
}




# Action de renommage
$btnRenommer.Add_Click({
    $fichierSelectionne = $comboFiles.SelectedItem
    $choix = $comboOptions.SelectedItem
    if (-not $fichierSelectionne -or -not $choix) {
        $labelStatus.Text = "Veuillez choisir un fichier et une action."
        return
    }

    $prefixe = switch -Regex ($choix) {
        "^1" { "1ere_relance" }
        "^2" { "2eme_relance" }
        "^3" { "3eme_relance" }
        "^4" { "4eme_relance" }
	"^5" { "Information_complementaire" }
	"^6" { "Demande_de_devis" }
        "^7" { "Devis" }
	"^8" { "Demande_d_intervention" }
        "^9" { "Rapport_d_intervention" }
        "^0" { "Facture" }
        "^R" { "Reponse" }
        default { "Fichier" }
    }

    # Génération de la date du jour au format yyyyMMdd
    $date = Get-Date -Format "yyyyMMdd"

    # Séparation du nom de fichier en nom sans extension et extension
    $nomSansExtension = [System.IO.Path]::GetFileNameWithoutExtension($fichierSelectionne)
    $extension = [System.IO.Path]::GetExtension($fichierSelectionne)
    $nouveauNom = "${date}_${prefixe}${extension}"

    try {
        Rename-Item -Path $fichierSelectionne -NewName $nouveauNom -ErrorAction Stop
        $labelStatus.ForeColor = 'DarkGreen'
	$labelStatus.MaximumSize = New-Object System.Drawing.Size(475, 0)  # Largeur max, hauteur auto
        $labelStatus.Font = New-Object System.Drawing.Font($labelStatus.Font, [System.Drawing.FontStyle]::Bold)
	$labelStatus.Text = "'$fichierSelectionne' s'appelle maintenant : '$nouveauNom'."
        # Mise à jour de la liste
        $comboFiles.Items.Clear()
        $script:msgFiles = Get-ChildItem -Path "." -Filter "*.msg" | Select-Object -ExpandProperty Name
        $comboFiles.Items.AddRange($msgFiles)
        if ($msgFiles.Count -gt 0) {
            $comboFiles.SelectedIndex = 0
        }
    } catch {
        $labelStatus.ForeColor = 'DarkRed'
        $labelStatus.Text = "Erreur : $_"
    }
})

# Lancement
[void]$form.ShowDialog()
