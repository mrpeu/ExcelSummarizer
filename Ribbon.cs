using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration.CustomUI;
using ExcelSummarizer.Properties;
using System.Windows.Forms;

namespace ExcelSummarizer
{
    [System.Runtime.InteropServices.ComVisible( true )]
    public class Ribbon : ExcelRibbon
    {
        public event EventHandler SettingsClicked;
        public event EventHandler SummaryClicked;
        public event EventHandler<EventArgs<string>> TemplatePathChanged;
        public event EventHandler<EventArgs<string>> TargetPathChanged;

        #region var
        public static IRibbonUI _ribbonUi { get; private set; }

        public enum EControlIds
        {
            grp_main,
            btn_summary, btn_settings,
            grp_configuration,
            lbl_template, txt_template, btn_template,
            lbl_target, txt_target, btn_target
        };

<<<<<<< HEAD
        static Dictionary<EControlIds, String> Labels = new Dictionary<EControlIds, String>() {
            {EControlIds.grp_main, "Erstellen"},
            {EControlIds.btn_summary, "Go"},
=======
        internal Dictionary<EControlIds, String> Labels = new Dictionary<EControlIds, String>() {
            {EControlIds.grp_main, "main"},
            {EControlIds.btn_summary, "Erstellen"},
>>>>>>> 5239de275f06cfc1c71131ec3c2410899efe91a9
            {EControlIds.btn_settings, "Einstellen"},

            {EControlIds.grp_configuration, "Ziele"},
            {EControlIds.btn_target, "Ordner ..."},
            
            {EControlIds.btn_template, "Vorlage: ..."}
        };

        internal Dictionary<EControlIds, String> Screentips = new Dictionary<EControlIds, String>() {
            {EControlIds.grp_main, String.Empty},
            {EControlIds.btn_summary, String.Empty},
            {EControlIds.btn_settings, String.Empty},
            {EControlIds.txt_template , String.Empty}
        };

        // rmk: "&#13;" for new line
        internal Dictionary<EControlIds, String> Supertips = new Dictionary<EControlIds, String>() {
            {EControlIds.grp_main, String.Empty},
            {EControlIds.btn_summary, String.Empty},
            {EControlIds.btn_settings, String.Empty},
            {EControlIds.txt_template , String.Empty},
            {EControlIds.txt_target, String.Empty}
        };

        internal Dictionary<EControlIds, Image> Images = new Dictionary<EControlIds, Image>() {
            {EControlIds.btn_summary, Resources.sum},
            {EControlIds.btn_settings, Resources.settings},
            {EControlIds.txt_template, null},
            {EControlIds.txt_target, null}
        };

        internal Dictionary<EControlIds, String> Texts = new Dictionary<EControlIds, String>() {
            {EControlIds.txt_template, String.Empty},
            {EControlIds.txt_target, String.Empty}
        };

        static Dictionary<EControlIds, bool> EnableStates = new Dictionary<EControlIds, bool>(){
            {EControlIds.btn_summary, true}
        };

        #region CustomUI
        // http://msdn.microsoft.com/en-us/library/vstudio/aa722523.aspx
        private const string CustomUI = @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'
            onLoad='Ribbon_Load'><ribbon><tabs><tab id='mainRibbonTab' label='Zusammenfassung' visible='true'>

            <group id='grp_configuration' getLabel='GetLabel'>

                <box id='big_box0' boxStyle='horizontal'>

                    <box id='big_box00' boxStyle='vertical'>

<<<<<<< HEAD
                        <button id='btn_target' size='normal' getLabel='GetLabel' getScreentip='GetScreentip' getSupertip='GetSupertip' onAction='btn_target_OnClick' />
=======
                        <labelControl id='lbl_spacer0' label=' ' />

                        <labelControl id='lbl_target' getLabel='GetLabel' />
>>>>>>> 5239de275f06cfc1c71131ec3c2410899efe91a9

                        <button id='btn_template' size='normal' getLabel='GetLabel' getScreentip='GetScreentip' getSupertip='GetSupertip' onAction='btn_template_OnClick' />

                    </box>


                    <box id='big_box01' boxStyle='vertical'>

<<<<<<< HEAD
=======
                        <labelControl id='lbl_spacer1' label=' ' />

                        <box id='box_target' boxStyle='horizontal'>
>>>>>>> 5239de275f06cfc1c71131ec3c2410899efe91a9

                        <box id='box_target' boxStyle='horizontal'>

                            <editBox id='txt_target' getImage='GetImage' showLabel='false' getText='GetText' onChange='OnTextChange' 
                                sizeString='WWWWWWWWWWWWWWWWWWWWWWWWWW'/>
                
                        </box>

                        <box id='box_template' boxStyle='horizontal'>

                            <editBox id='txt_template' getImage='GetImage' showLabel='false' getText='GetText' onChange='OnTextChange' 
                                sizeString='WWWWWWWWWWWWWWWWWWWWWWWWWW'/>
                
                        </box>

                    </box>

                </box>

            </group>

            <group id='grp_main' getLabel='GetLabel'>
<!--
                <button id='btn_settings' size='large' getImage='GetImage' getLabel='GetLabel' getScreentip='GetScreentip' getSupertip='GetSupertip' onAction='btn_settings_OnClick' />
-->
                <button id='btn_summary' size='large' getImage='GetImage' getLabel='GetLabel' getEnabled='GetEnabledState' getScreentip='GetScreentip' getSupertip='GetSupertip' onAction='btn_summary_OnClick' />

            </group>

            </tab></tabs></ribbon>
        </customUI >";
        #endregion

        #endregion


        #region init
        public override string GetCustomUI( string uiName )
        {
            // todo idea: parse the xml and use the ids directly instead of this.EControlIDs
            // see GetCustomUI https://github.com/brymck/finansu/blob/master/FinAnSu/Controls/Ribbon.cs

            return CustomUI;
        }

        public void Ribbon_Load( IRibbonUI sender )
        {
            _ribbonUi = sender;
            AddIn.RegisterRibbon( this );
        }

        public override void OnStartupComplete( ref Array custom )
        {
            base.OnStartupComplete( ref custom );
        }
        #endregion


        public String GetLabel( IRibbonControl control )
        {
            String label = String.Empty;

            EControlIds id;
            if ( !Enum.TryParse<EControlIds>( control.Id, out id ) ) throw new Exception( "Incorrect RibbonControl. Unknown id: " + control.Id );

            if ( !Labels.TryGetValue( id, out label ) )
            {
                label = "<Error>";
            }

            return label;
        }

        public string GetScreentip( IRibbonControl control )
        {
            string screentip = String.Empty;

            EControlIds id;
            if ( !Enum.TryParse<EControlIds>( control.Id, out id ) ) throw new Exception( "Incorrect RibbonControl. Unknown id: " + control.Id );

            if ( !Screentips.TryGetValue( id, out screentip ) )
            {
                screentip = String.Empty;
            }

            return screentip;
        }

        public string GetSupertip( IRibbonControl control )
        {
            string supertip = String.Empty;

            EControlIds id;
            if ( !Enum.TryParse<EControlIds>( control.Id, out id ) ) throw new Exception( "Incorrect RibbonControl. Unknown id: " + control.Id );

            if ( !Supertips.TryGetValue( id, out supertip ) )
            {
                supertip = String.Empty;
            }

            return supertip;
        }

        public Image GetImage( IRibbonControl control )
        {
            Image img = null;

            EControlIds id;
            if ( !Enum.TryParse<EControlIds>( control.Id, out id ) ) throw new Exception( "Incorrect RibbonControl. Unknown id: " + control.Id );

            Images.TryGetValue( id, out img );

            return img;
        }

        public String GetText( IRibbonControl control )
        {
            String text = null;

            EControlIds id;
            if ( !Enum.TryParse<EControlIds>( control.Id, out id ) ) throw new Exception( "Incorrect RibbonControl. Unknown id: " + control.Id );

            Texts.TryGetValue( id, out text );

            return text;
        }

        public bool GetEnabledState( IRibbonControl control )
        {
            EControlIds id;
            if ( !Enum.TryParse<EControlIds>( control.Id, out id ) ) throw new Exception( "Incorrect RibbonControl. Unknown id: " + control.Id );

            return EnableStates.ContainsKey( id ) && EnableStates[ id ];
        }


        public void btn_target_OnClick( IRibbonControl control )
        {
            var dialog = new FolderBrowserDialog();
            var id = EControlIds.txt_target;
            dialog.SelectedPath = Texts[ id ];

            if ( dialog.ShowDialog() == DialogResult.OK )
            {
                Texts[ id ] = dialog.SelectedPath;
                OnTextChange( control, Texts[ id ] );
                _ribbonUi.InvalidateControl( id.ToString() );
            }
        }

        public void btn_template_OnClick( IRibbonControl control )
        {
            var dialog = new OpenFileDialog();
            var id = EControlIds.txt_template;
            dialog.InitialDirectory = System.IO.Path.GetFullPath( System.IO.Path.GetDirectoryName( Texts[ id ] ) );
            dialog.FileName = System.IO.Path.GetFileName( Texts[ id ] );

            if ( dialog.ShowDialog() == DialogResult.OK )
            {
                Texts[ id ] = System.IO.Path.GetFullPath( dialog.FileName );
                OnTextChange( control, Texts[ id ] );
                _ribbonUi.InvalidateControl( id.ToString() );
            }
        }

        public void btn_settings_OnClick( IRibbonControl control )
        {
            if ( SettingsClicked != null ) SettingsClicked( control, EventArgs.Empty );
        }

        public void btn_summary_OnClick( IRibbonControl control )
        {
            if ( SummaryClicked != null ) SummaryClicked( control, EventArgs.Empty );
        }

        public void OnTextChange( IRibbonControl control, String text )
        {
            EControlIds id;
            if ( !Enum.TryParse<EControlIds>( control.Id, out id ) ) throw new Exception( "Incorrect RibbonControl. Unknown id: " + control.Id );

            switch ( id )
            {
                case EControlIds.btn_template:
                case EControlIds.txt_template:
                    if ( TemplatePathChanged != null ) TemplatePathChanged( this, new EventArgs<string>( text ) );
                    break;

                case EControlIds.btn_target:
                case EControlIds.txt_target:
                    if ( TargetPathChanged != null ) TargetPathChanged( this, new EventArgs<string>( text ) );
                    break;

                default:
                    break;
            }

        }

        internal void Invalidate()
        {
<<<<<<< HEAD
            Texts[ EControlIds.txt_template ] = Configuration.TemplatePath;
            Images[ EControlIds.txt_template ] = Configuration.IsTemplateValid ? Resources.bullet_green : Resources.bullet_pink;

            Texts[ EControlIds.txt_target ] = Configuration.TargetPath;
            Images[ EControlIds.txt_target ] = Configuration.IsTargetValid ? Resources.bullet_green : Resources.bullet_pink;

            EnableStates[ EControlIds.btn_summary ] = Configuration.IsTemplateValid && Configuration.IsTargetValid;

=======
>>>>>>> 5239de275f06cfc1c71131ec3c2410899efe91a9
            _ribbonUi.Invalidate();
        }
    }
}
