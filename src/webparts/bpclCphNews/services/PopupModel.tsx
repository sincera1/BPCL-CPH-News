import * as React from "react";
import { Button } from "react-bootstrap";

export type PopupType = "success" | "danger" | "confirm";

interface PopupModalProps {
  show: boolean;
  type: PopupType;
  message: string;
  onConfirm?: () => void;
  onCancel?: () => void;
  onClose?: (result?: "yes" | "no") => void;
}

const PopupModal: React.FC<PopupModalProps> = ({
  show,
  type,
  message,
  onConfirm,
  onCancel,
  onClose,
}) => {
  if (!show) return null;

  const isConfirm = type === "confirm";

  return (
    <div
      className="modal fade show"
      style={{ display: "block", backgroundColor: "rgba(0,0,0,0.5)" }}
      role="dialog"
      aria-modal="true"
    >
      <div className="modal-dialog modal-dialog-centered" role="document">
        <div className="modal-content">
          <div className="modal-header">
            <h5 className="modal-title">
              {isConfirm ? "Confirmation" : "Status"}
            </h5>
          </div>

          <div className="modal-body">
            <p style={{ whiteSpace: "pre-wrap", margin: 0 }}>{message}</p>
          </div>

          <div className="modal-footer">
            {isConfirm ? (
              <>
                {/* NO button */}
                <Button
                  variant="secondary"
                  onClick={() => onClose && onClose("no")}
                >
                  No
                </Button>

                {/* YES button */}
                <Button
                  variant="danger"
                  onClick={() => {
                    if (onConfirm) onConfirm();
                    if (onClose) onClose("yes");
                  }}
                >
                  Yes
                </Button>
              </>
            ) : (
              <Button
                className="btn"
                style={{
                  minWidth: "fit-content",
                  height: "40px",
                  backgroundColor: "#013396",
                  border: "1px solid #013396",
                  paddingLeft: "1.6rem",
                  paddingRight: "1.6rem",
                }}
                onClick={() => onClose && onClose()}
              >
                OK
              </Button>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

export default PopupModal;
