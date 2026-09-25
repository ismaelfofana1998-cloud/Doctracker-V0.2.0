using System;
using System.Drawing;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Doctracker.AddIn.UI;
using Doctracker.Core.Models;
using Microsoft.Office.Core;

namespace Doctracker.AddIn.Ribbon
{
    [ComVisible(true)]
    public sealed class DoctrackerRibbon : IRibbonExtensibility
    {
        private const string RibbonXml = @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnLoad'>
  <ribbon>
    <tabs>
      <tab id='DoctrackerTab' label='Doctracker'>
        <group id='ProjectGroup' label='Dossier'>
          <button id='OpenPane' getImage='GetImage' label='Ouvrir Doctracker' size='large' onAction='OpenPane_OnAction'/>
          <button id='ImportDocuments' getImage='GetImage' label='Importer' size='large' onAction='ImportDocuments_OnAction'/>
          <button id='CrossReference' getImage='GetImage' label='Créer / modifier Xref' size='large' onAction='Command_OnAction'/>
          <button id='ExportPdf' getImage='GetImage' label='Exporter' size='large' onAction='Command_OnAction'/>
          <menu id='DocumentsMenu' label='Documents' getImage='GetImage'>
            <button id='ImportFolder' label='Importer un dossier…' onAction='Command_OnAction'/>
            <button id='Categorize' label='Classer les documents…' onAction='Command_OnAction'/>
            <button id='TestReference' label='Référence du test…' onAction='Command_OnAction'/>
            <button id='Reindex' label='Reconnaître le dossier sélectionné' onAction='Command_OnAction'/>
            <button id='ReindexOcr' label='Réindexer le document actif par OCR' onAction='Command_OnAction'/>
            <button id='Remove' label='Retirer le document actif' onAction='Command_OnAction'/>
          </menu>
          <menu id='RecoveryMenu' label='Récupération' getImage='GetImage'>
            <button id='Backup' label='Sauvegarder les pièces et liens…' onAction='Command_OnAction'/>
            <button id='Restore' label='Restaurer une sauvegarde…' onAction='Command_OnAction'/>
            <button id='RepairLinks' label='Réparer les liens' onAction='Command_OnAction'/>
            <button id='Diagnostics' label='Diagnostic des erreurs' onAction='Command_OnAction'/>
            <button id='RecoveryFolder' label='Afficher le cache local' onAction='Command_OnAction'/>
            <button id='SharedStorage' label='Stockage autonome / partagé…' onAction='Command_OnAction'/>
          </menu>
          <button id='SearchDocuments'  getImage='GetImage' label='Rechercher la cellule' onAction='Search_OnAction'/>
        </group>
        <group id='SnipGroup' label='Snips'>
          <toggleButton id='TextSnip' getImage='GetImage' label='Texte' onAction='TextSnip_OnAction' getPressed='TextSnip_GetPressed' size='large'/>
          <toggleButton id='ValidationSnip' getImage='GetImage' label='Validation' onAction='ValidationSnip_OnAction' getPressed='ValidationSnip_GetPressed' size='large'/>
          <toggleButton id='ExceptionSnip' getImage='GetImage' label='Exception' onAction='ExceptionSnip_OnAction' getPressed='ExceptionSnip_GetPressed' size='large'/>
          <toggleButton id='NumberSnip' getImage='GetImage' label='Nombre' onAction='NumberSnip_OnAction' getPressed='NumberSnip_GetPressed' size='large'/>
          <toggleButton id='DateSnip' getImage='GetImage' label='Date' onAction='DateSnip_OnAction' getPressed='DateSnip_GetPressed' size='large'/>
          <toggleButton id='SumSnip' getImage='GetImage' label='Somme' onAction='SumSnip_OnAction' getPressed='SumSnip_GetPressed' size='large'/>
          <toggleButton id='TableSnip' getImage='GetImage' label='Tableau' onAction='TableSnip_OnAction' getPressed='TableSnip_GetPressed' size='large'/>
          <toggleButton id='Comment' getImage='GetImage' label='Commentaire' size='large' onAction='Comment_OnAction' getPressed='Comment_GetPressed'/>
          <button id='DeleteSelectionSnips' label='Supprimer les snips de la plage' onAction='Command_OnAction'/>
          <button id='DeleteSnip' getImage='GetImage' label='Supprimer un snip' onAction='Command_OnAction'/>
        </group>
        <group id='MatchingGroup'  label='Contrôle'>
          <button id='SetMatchInput'  getImage='GetImage' label='Définir recherche' onAction='SetMatchInput_OnAction'/>
          <button id='SetMatchOutput' getImage='GetImage' label='Définir résultat' onAction='SetMatchOutput_OnAction'/>
          <button id='Match' getImage='GetImage' label='Lancer le matching' size='large' onAction='Match_OnAction'/>
          <button id='OpenProof' getImage='GetImage' label='Ouvrir la preuve' onAction='OpenProof_OnAction'/>
          <button id='ReviewProof' getImage='GetImage' label='Revoir' onAction='ReviewProof_OnAction'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        private readonly Dictionary<string, object> images = new Dictionary<string, object>();
        public object GetImage(IRibbonControl control)
        {
            object image;
            if (images.TryGetValue(control.Id, out image)) return image;
            var key = control.Id.EndsWith("Snip", StringComparison.Ordinal) ? control.Id.Substring(0, control.Id.Length - 4) : control.Id;
            SnipType type;
            using (var bitmap = SnipTheme.Icon(key, 32, Enum.TryParse(key, out type) ? SnipTheme.ColorFor(type) : SnipTheme.Ink))
                image = PictureConverter.Convert(bitmap);
            images.Add(control.Id, image);
            return image;
        }
        private sealed class PictureConverter : AxHost
        {
            private PictureConverter() : base("") { }
            public static object Convert(Image image) => GetIPictureDispFromPicture(image);
        }

        private IRibbonUI ribbon;

        public string GetCustomUI(string ribbonId)
        {
            return RibbonXml;
        }

        internal static DoctrackerRibbon Instance { get; private set; }
        internal void Refresh() => ribbon?.Invalidate();

        public void OnLoad(IRibbonUI ribbonUi)
        {
            ribbon = ribbonUi;
            Instance = this;
        }

        public void OpenPane_OnAction(IRibbonControl control) => Controller.Toggle();
        public void ImportDocuments_OnAction(IRibbonControl control) => Controller.ImportDocuments();
        public void Command_OnAction(IRibbonControl control) { Controller.ExecuteCommand(control.Id); Refresh(); }
        public void Comment_OnAction(IRibbonControl control,bool pressed) { Controller.ExecuteCommand("Comment"); Refresh(); }
        public bool Comment_GetPressed(IRibbonControl control) => Globals.ThisAddIn?.Controller?.CommandPressed("Comment") ?? false;
        public void Search_OnAction(IRibbonControl control) => Controller.SearchSelection();

        public void ValidationSnip_OnAction(IRibbonControl control, bool pressed) => SetSnipMode(SnipType.Validation, pressed);
        public void ExceptionSnip_OnAction(IRibbonControl control, bool pressed) => SetSnipMode(SnipType.Exception, pressed);
        public void TextSnip_OnAction(IRibbonControl control, bool pressed) => SetSnipMode(SnipType.Text, pressed);
        public void NumberSnip_OnAction(IRibbonControl control, bool pressed) => SetSnipMode(SnipType.Number, pressed);
        public void DateSnip_OnAction(IRibbonControl control, bool pressed) => SetSnipMode(SnipType.Date, pressed);
        public void SumSnip_OnAction(IRibbonControl control, bool pressed) => SetSnipMode(SnipType.Sum, pressed);
        public void TableSnip_OnAction(IRibbonControl control, bool pressed) => SetSnipMode(SnipType.Table, pressed);

        public bool ValidationSnip_GetPressed(IRibbonControl control) => Globals.ThisAddIn?.Controller?.IsSnipMode(SnipType.Validation) ?? false;
        public bool ExceptionSnip_GetPressed(IRibbonControl control) => Globals.ThisAddIn?.Controller?.IsSnipMode(SnipType.Exception) ?? false;
        public bool TextSnip_GetPressed(IRibbonControl control) => Globals.ThisAddIn?.Controller?.IsSnipMode(SnipType.Text) ?? false;
        public bool NumberSnip_GetPressed(IRibbonControl control) => Globals.ThisAddIn?.Controller?.IsSnipMode(SnipType.Number) ?? false;
        public bool DateSnip_GetPressed(IRibbonControl control) => Globals.ThisAddIn?.Controller?.IsSnipMode(SnipType.Date) ?? false;
        public bool SumSnip_GetPressed(IRibbonControl control) => Globals.ThisAddIn?.Controller?.IsSnipMode(SnipType.Sum) ?? false;
        public bool TableSnip_GetPressed(IRibbonControl control) => Globals.ThisAddIn?.Controller?.IsSnipMode(SnipType.Table) ?? false;

        public void SetMatchInput_OnAction(IRibbonControl control) => Controller.SetMatchingInputSelection();
        public void SetMatchOutput_OnAction(IRibbonControl control) => Controller.SetMatchingOutputSelection();
        public void Match_OnAction(IRibbonControl control) => Controller.MatchSelection();
        public void OpenProof_OnAction(IRibbonControl control) => Controller.NavigateFromSelection();
        public void ReviewProof_OnAction(IRibbonControl control) => Controller.ReviewSelection();

        private void SetSnipMode(SnipType type, bool pressed)
        {
            Controller.SetSnipMode(pressed ? type : (SnipType?)null);
            ribbon?.InvalidateControl("Comment");
            ribbon?.InvalidateControl("ValidationSnip");
            ribbon?.InvalidateControl("ExceptionSnip");
            ribbon?.InvalidateControl("TextSnip");
            ribbon?.InvalidateControl("NumberSnip");
            ribbon?.InvalidateControl("DateSnip");
            ribbon?.InvalidateControl("SumSnip");
            ribbon?.InvalidateControl("TableSnip");
        }

        private static PaneController Controller
        {
            get
            {
                var controller = Globals.ThisAddIn?.Controller;
                if (controller == null)
                    throw new InvalidOperationException("Doctracker is not initialized.");
                return controller;
            }
        }
    }
}
