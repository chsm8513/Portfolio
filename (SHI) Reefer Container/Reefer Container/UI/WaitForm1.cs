using DevExpress.Utils.Extensions;
using DevExpress.XtraWaitForm;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Reefer_Container
{
    public partial class WaitForm1 : WaitForm
    {
        public WaitForm1()
        {
            InitializeComponent();
            this.progressPanel1.AutoHeight = true;
        }

        #region Overrides

        public override void SetCaption(string caption)
        {
            base.SetCaption(caption);
            this.progressPanel1.Caption = caption;
        }
        public override void SetDescription(string description)
        {
            base.SetDescription(description);
            this.progressPanel1.Description = description;
        }
        public override void ProcessCommand(Enum cmd, object arg)
        {
            base.ProcessCommand(cmd, arg);
            // Form1에서 보낸 'Bounds' 정보(arg)가 Rectangle 타입인지 확인합니다.
            if (arg is Rectangle parentBounds)
            {
                // 맞다면, 해당 위치를 기준으로 폼을 중앙에 배치하는 메서드를 호출합니다.
                CenterToParent(parentBounds);
            }
        }
        // CenterToParent 헬퍼 메서드를 추가합니다.
        private void CenterToParent(Rectangle parentBounds)
        {
            // 부모 폼의 중앙 좌표 계산
            int parentCenterX = parentBounds.X + (parentBounds.Width / 2);
            int parentCenterY = parentBounds.Y + (parentBounds.Height / 2);

            // WaitForm의 새 위치 계산
            int newX = parentCenterX - (this.Width / 2);
            int newY = parentCenterY - (this.Height / 2);

            // 계산된 위치로 이동
            this.Location = new Point(newX, newY);
        }
        #endregion

    }
}
