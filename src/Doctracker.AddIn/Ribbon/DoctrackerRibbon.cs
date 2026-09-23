using System;
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
          <button id='OpenPane' label='Ouvrir Doctracker' size='large' onAction='OpenPane_OnAction'/>
          <button id='ImportDocuments' label='Ajouter des pièces' onAction='ImportDocuments_OnAction'/>
          <button id='SearchDocuments' label='Rechercher' size='large' onAction='Search_OnAction'/>
        </group>
        <group id='SnipGroup' label='Mode snip — actif jusqu’à désactivation'>
          <toggleButton id='ValidationSnip' label='Validation' onAction='ValidationSnip_OnAction' getPressed='ValidationSnip_GetPressed'/>
          <toggleButton id='ExceptionSnip' label='Exception' onAction='ExceptionSnip_OnAction' getPressed='ExceptionSnip_GetPressed'/>
          <toggleButton id='TextSnip' label='Texte' onAction='TextSnip_OnAction' getPressed='TextSnip_GetPressed'/>
          <toggleButton id='NumberSnip' label='Nombre' onAction='NumberSnip_OnAction' getPressed='NumberSnip_GetPressed'/>
          <toggleButton id='DateSnip' label='Date' onAction='DateSnip_OnAction' getPressed='DateSnip_GetPressed'/>
          <toggleButton id='SumSnip' label='Somme' onAction='SumSnip_OnAction' getPressed='SumSnip_GetPressed'/>
          <toggleButton id='TableSnip' label='Tableau' onAction='TableSnip_OnAction' getPressed='TableSnip_GetPressed'/>
        </group>
        <group id='MatchingGroup' label='Contrôle'>
          <button id='SetMatchInput' label='Définir recherche' onAction='SetMatchInput_OnAction'/>
          <button id='SetMatchOutput' label='Définir résultat' onAction='SetMatchOutput_OnAction'/>
          <button id='Match' label='Lancer le matching' size='large' onAction='Match_OnAction'/>
          <button id='OpenProof' label='Ouvrir la preuve' onAction='OpenProof_OnAction'/>
          <button id='ReviewProof' label='Revoir' onAction='ReviewProof_OnAction'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

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
