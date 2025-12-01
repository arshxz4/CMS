/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @rushstack/no-new-null */
/* eslint-disable no-void */
/* eslint-disable @typescript-eslint/no-unused-expressions */
/* eslint-disable no-unused-expressions*/
/* eslint-disable  prefer-const */
/* eslint-disable  react/no-unescaped-entities */
/*eslint-disable @typescript-eslint/no-use-before-define */
//iosthreiht
import * as React from "react";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
// import { faPlus, faTrash } from "@fortawesome/free-solid-svg-icons";
import "./RequesterInvoiceSection.module.scss";
import {
  updateDataToSharePoint,
  getSharePointData,
} from "../services/SharePointService"; // Adjust the import path as necessary
import { Modal, Button } from "react-bootstrap"; // Import Bootstrap Modal
import FinaceInvoiceSection from "./FinaceInvoiceSection";
import moment from "moment";
import { DatePicker } from "antd";
import {
  getDocumentLibraryDataWithSelect,
  handleDownload,
} from "../services/SharePointService";
import {
  // faPlus,
  faTrash,
  faClockRotateLeft,
  // faXmark
  // faAngleUp,
  // faAngleDown,
} from "@fortawesome/free-solid-svg-icons";
import CreditNoteDetails from "./CreditNoteDetails";
interface InvoiceRow {
  id: number;
  InvoiceDescription: string;
  RemainingPoAmount: string;
  InvoiceAmount: string;
  InvoiceDueDate: string;
  InvoiceProceedDate: string;
  showProceed: boolean;
  InvoiceStatus: string; // Add InvoiceStatus to the interface
  PrevInvoiceStatus?: string; // Add PrevInvoiceStatus to track previous status
  CreditNoteStatus?: string; // Add CreditNoteStatus to track credit note status
  userInGroup: boolean; // Add userInGroup to the interface
  employeeEmail: string; // Add employeeEmail to the interface
  itemID: number | null;
  InvoiceNo: string;
  InvoiceDate: string;
  InvoiceTaxAmount: string;
  ClaimNo: number | null;
  DocId: string;
  InvoiceFileID: string;
  invoiceApprovalChecked?: boolean; // Add this property
  closeInvoiceChecked?: boolean;
  CloseAmount?: string;
  CloseReason?: string;
}

export default function RequesterInvoiceSection({
  userGroups,
  invoiceRows,
  setInvoiceRows,
  handleInvoiceChange,
  addInvoiceRow,
  totalPoAmount,
  errors,
  isEditMode, // Add a prop to indicate edit mode
  approverStatus,
  currentUserEmail, // Add currentUser as a prop
  siteUrl,
  context, // Add context as a prop
  props,
  hideAddInvoiceButton,
  poAmount,
  startDate,
  endDate, // New prop
  disableDeleteInvoiceRow, // <-- New prop
  onProceedButtonCountChange,
  isCollapsed,
  setIsCollapsed, // Add to props:
  selectedRowDetails,
  // new props for approval checkbox (optional)
  invoiceApprovalChecked,
  invoiceCloseApprovalChecked,
  setInvoiceApprovalChecked,
  setDeletedInvoiceItemIDs,
}: {
  userGroups: any;
  invoiceRows: InvoiceRow[];
  setInvoiceRows: React.Dispatch<React.SetStateAction<InvoiceRow[]>>;
  handleInvoiceChange: (
    index: number,
    field: string,
    value: string | number
  ) => void;

  addInvoiceRow: () => void;
  totalPoAmount: number;
  errors: { [key: string]: string };
  isEditMode: boolean;
  approverStatus: string;
  currentUserEmail: string;
  siteUrl: string;
  context: any;
  props: any;
  hideAddInvoiceButton: boolean;
  poAmount: number;
  startDate: any;
  endDate: any;
  disableDeleteInvoiceRow?: boolean;
  onProceedButtonCountChange?: (count: number) => void;

  isCollapsed: boolean;
  setIsCollapsed: React.Dispatch<React.SetStateAction<boolean>>;
  selectedRowDetails?: any;
  invoiceApprovalChecked?: boolean;
  invoiceCloseApprovalChecked?: boolean;
  setInvoiceApprovalChecked?: React.Dispatch<React.SetStateAction<boolean>>;
  setDeletedInvoiceItemIDs?: React.Dispatch<React.SetStateAction<number[]>>;
}) {
  const InvoiceList = "CMSRequestDetails";
  const InvoiceHistory = "CMSPaymentHistory";
  // console.log(siteUrl, "siteUrlinvoice123");
  // console.log(props, "propsinvoice123");
  const [showEditModal, setShowEditModal] = React.useState(false); // State to control modal visibility
  const [selectedRow, setSelectedRow] = React.useState<InvoiceRow | null>(null); // State to store selected row data
  const [invoiceHistoryData, setInvoiceHistoryData] = React.useState<any[]>([]);
  const [showHistoryModal, setShowHistoryModal] = React.useState(false);
  const [historyLoading, setHistoryLoading] = React.useState(false);
  const [invoiceDocuments, setInvoiceDocuments] = React.useState<any[]>([]);
  const [showEditInvoiceColumn, setShowEditInvoiceColumn] =
    React.useState(false); // Add a state to track the visibility of the "Edit Invoice" column

  const [localInvoiceData, setLocalInvoiceData] = React.useState(
    invoiceRows.map((row) => ({
      InvoiceDescription: row.InvoiceDescription,
      InvoiceAmount: row.InvoiceAmount,
      InvoiceDueDate: row.InvoiceDueDate,
      RemainingPoAmount: row.RemainingPoAmount,
    }))
  );

const [showCloseSection, setShowCloseSection] = React.useState(false);
const [closeAmount, setCloseAmount] = React.useState("");
const [closeReason, setCloseReason] = React.useState("");
const [autoSelectedCloseRows, setAutoSelectedCloseRows] = React.useState(false);
const [closeFieldsLoaded, setCloseFieldsLoaded] = React.useState(false);
const requestorEmail = props.selectedRow?.employeeEmail;
const managerEmail = props.selectedRow?.projectMangerEmail;
const isRequestorManager = requestorEmail === managerEmail;
const isManagerButNotRequester =
  currentUserEmail?.toLowerCase() === managerEmail?.toLowerCase() &&
  requestorEmail?.toLowerCase() !== managerEmail?.toLowerCase();

const [showCloseRejectModal, setShowCloseRejectModal] = React.useState(false);
const [showCloseHoldModal, setShowCloseHoldModal] = React.useState(false);

const [managerCloseReason, setManagerCloseReason] = React.useState("");


// keep selected rows in sync with close inputs
React.useEffect(() => {
  // Requester writes values
  if (currentUserEmail === requestorEmail) {
    setInvoiceRows((prev) =>
      prev.map((r) =>
        r.closeInvoiceChecked
          ? { ...r, CloseAmount: closeAmount, CloseReason: closeReason }
          : r
      )
    );
  }
}, [closeAmount, closeReason, currentUserEmail, requestorEmail]);


React.useEffect(() => {
  if (
    !autoSelectedCloseRows &&                              
    managerEmail &&
    currentUserEmail.toLowerCase() === managerEmail.toLowerCase()
  ) {
    setInvoiceRows(prev =>
      prev.map(r =>
        r.InvoiceStatus === "Pending Close Approval"
          ? { ...r, closeInvoiceChecked: true }
          : r
      )
    );

    setAutoSelectedCloseRows(true);    // run ONLY once
  }
}, [autoSelectedCloseRows, currentUserEmail, managerEmail]);



React.useEffect(() => {
  // Make sure emails exist
  if (!managerEmail || !currentUserEmail) return;

  // Only manager
  if (currentUserEmail.toLowerCase() !== managerEmail.toLowerCase()) return;

  // Invoice rows not ready yet
  if (!invoiceRows || invoiceRows.length === 0) return;

  // Already fetched once – do not run again
  if (closeFieldsLoaded) return;

  const fetchCloseFields = async () => {
    console.log("Fetching manager close fields...");

    let didUpdate = false;

    for (const row of invoiceRows) {
      if (!row.itemID) continue;

      const resp = await getSharePointData(
        { context: props.context },
        InvoiceList,
        `$filter=Id eq ${row.itemID}`
      );

      if (resp && resp.length > 0) {
        const item = resp[0];

        didUpdate = true;

        setInvoiceRows(prev =>
          prev.map(r =>
            r.itemID === row.itemID
              ? {
                  ...r,
                  CloseAmount: item.CloseAmount,
                  CloseReason: item.CloseReason,
                }
              : r
          )
        );
      }
    }

    if (didUpdate) {
      setCloseFieldsLoaded(true);    // <-- IMPORTANT
    }
  };

  void fetchCloseFields();
}, [
  invoiceRows.length, // <-- depends ONLY on length
  managerEmail,
  currentUserEmail,
  closeFieldsLoaded
]);


const getPendingCloseInvoices = () => {
  return invoiceRows.filter(
    (row) => row.InvoiceStatus === "Pending Close Approval"
  );
};

const refreshInvoices = async () => {
  try {
    const requestId = props.selectedRow?.RequestID;

    if (!requestId) {
      console.error("Missing RequestID — cannot refresh invoices");
      return;
    }

    const filterQuery = `$filter=RequestID eq ${requestId}`;

    const resp = await getSharePointData(
      { context: props.context },
      InvoiceList,
      filterQuery
    );

    setInvoiceRows(resp);
  } catch (err) {
    console.error("Error refreshing invoices:", err);
  }
};



// const refreshInvoiceForHOLD = async () => {
//   try {

//     let requestId =
//       props.selectedRow?.RequestID ||
//       invoiceRows[0]?.RequestID ||
//       null;

//     if (!requestId) {
//       console.error("Missing RequestID — cannot refresh invoices");
//       return;
//     }

//     const filterQuery = `$filter=RequestID eq ${requestId}`;

//     const resp = await getSharePointData(
//       { context: props.context },
//       InvoiceList,
//       filterQuery
//     );

//     setInvoiceRows(resp);

//   } catch (err) {
//     console.error("Error refreshing invoices:", err);
//   }
// };



  // === MANAGER MODAL STATES ===
const [showRejectModal, setShowRejectModal] = React.useState(false);
const [showHoldModal, setShowHoldModal] = React.useState(false);
const [modalRow, setModalRow] = React.useState<InvoiceRow | null>(null);
const [managerReason, setManagerReason] = React.useState("");
const [managerDueDate, setManagerDueDate] = React.useState<string | null>(null);

// === OPEN / CLOSE MODALS ===
const openRejectModal = (row: InvoiceRow) => {
  setModalRow(row);
  setManagerReason("");
  setManagerDueDate(null);
  setShowRejectModal(true);
};

const openHoldModal = (row: InvoiceRow) => {
  setModalRow(row);
  setManagerReason("");
  setManagerDueDate(null);
  setShowHoldModal(true);
};

const closeModals = () => {
  setShowRejectModal(false);
  setShowHoldModal(false);
  setModalRow(null);
};

// === SUBMIT MANAGER DECISIONS ===
const submitReject = async () => {
  if (!modalRow) return;

  if (!managerReason.trim()) {
    alert("Reason is mandatory for Reject.");
    return;
  }

  await handleManagerDecision(
    modalRow,
    "Reject",
    managerDueDate ? moment(managerDueDate).format("YYYY-MM-DD") : "",
    managerReason
  );

  closeModals();
};


const submitHold = async () => {
  if (!modalRow) return;

  if (!managerReason.trim()) {
    alert("Reason is mandatory for Hold.");
    return;
  }

  await handleManagerDecision(
    modalRow,
    "Hold",
    "",
    managerReason
  );

  closeModals();
};





  const fetchAllInvoiceDocuments = async (siteUrl: string) => {
    const selectFields = "Id, FileLeafRef, FileRef,EncodedAbsUrl,DocID";
    const libraryName = "InvoiceDocument";
    // const filterQuery = `$top=5000`;
    const filterQuery = `$top=5000`;
    // console.log(siteUrl, "siteUrlinvoice");

    try {
      const response = await getDocumentLibraryDataWithSelect(
        libraryName,
        filterQuery,
        selectFields,
        siteUrl
      );
      // console.log("All INVOICE items:", response);
      return response;
    } catch (error) {
      console.error("Error fetching invoice Document items:", error);
      return [];
    }
  };
  // ...existing code...

  React.useEffect(() => {
    void (async () => {
      const docs = await fetchAllInvoiceDocuments(props.siteUrl);
      setInvoiceDocuments(docs);
    })();
  }, [props.siteUrl]);

  // const pendingStatuses = ["Hold", "Open", "Pending From Approver", "Reminder"];
  // const [proceedClicked, setProceedClicked] = React.useState(false);
  const [proceededRows, setProceededRows] = React.useState<number[]>([]);
  // const deleteInvoiceRow = (id: number) => {
  //   setInvoiceRows(invoiceRows.filter((row) => row.id !== id));
  // };
  /*
  const deleteInvoiceRow = (id: number) => {
    // find the row about to be deleted
    const rowToDelete = invoiceRows.find((row) => row.id === id);

    // if in edit mode and row has an existing itemID, push it into parent's deleted IDs array
    if (props?.rowEdit === "Yes" && rowToDelete?.itemID) {
      const numericItemId = Number(rowToDelete.itemID);
      if (
        !isNaN(numericItemId) &&
        typeof setDeletedInvoiceItemIDs === "function"
      ) {
        setDeletedInvoiceItemIDs((prev) => {
          // avoid duplicates
          if (prev.includes(numericItemId)) return prev;
          return [...prev, numericItemId];
        });
      }
    }
    if (props.rowEdit === "Yes") {
      syncUploadedCreditNoteRows();
    }

    // remove the row from UI
    setInvoiceRows(invoiceRows.filter((row) => row.id !== id));
  };
*/

  const deleteInvoiceRow = (id: number) => {
    // find the row about to be deleted
    const rowToDelete = invoiceRows.find((row) => row.id === id);

    // if in edit mode and row has an existing itemID, push it into parent's deleted IDs array
    if (props?.rowEdit === "Yes" && rowToDelete?.itemID) {
      const numericItemId = Number(rowToDelete.itemID);
      if (
        !isNaN(numericItemId) &&
        typeof setDeletedInvoiceItemIDs === "function"
      ) {
        setDeletedInvoiceItemIDs((prev) => {
          // avoid duplicates
          if (prev.includes(numericItemId)) return prev;
          return [...prev, numericItemId];
        });
      }
    }

    // Remove the row from UI
    let updatedRows = invoiceRows.filter((row) => row.id !== id);

    // If no rows left, add a blank row
    if (updatedRows.length === 0) {
      updatedRows = [
        {
          id: 1,
          InvoiceDescription: "",
          RemainingPoAmount: totalPoAmount.toFixed(2),
          InvoiceAmount: "",
          InvoiceDueDate: "",
          InvoiceProceedDate: "",
          showProceed: false,
          InvoiceStatus: "",
          userInGroup: false,
          employeeEmail: "",
          itemID: null,
          InvoiceNo: "",
          InvoiceDate: "",
          InvoiceTaxAmount: "",
          ClaimNo: null,
          DocId: "",
          InvoiceFileID: "",
          invoiceApprovalChecked: false,
          invoiceCloseApprovalChecked: false, // Initialize here
        },
      ];
    }

    setInvoiceRows(updatedRows);

    // Sync local fields after deletion
    if (props.rowEdit === "Yes") {
      setTimeout(() => {
        syncUploadedCreditNoteRows();
      }, 0);
    }
  };
  const handleTextFieldChange = (
    index: number,
    field: keyof InvoiceRow,
    value: string | number
  ) => {
    setInvoiceRows((prevRows) => {
      let updatedRows = [...prevRows];

      // Ensure row exists before updating
      if (!updatedRows[index]) {
        console.error(`Row at index ${index} is undefined.`);
        return prevRows;
      }

      // Update specific field in the row
      updatedRows[index] = { ...updatedRows[index], [field]: value };

      if (field === "InvoiceAmount") {
        const poAmt = parseFloat(totalPoAmount.toString()) || 0;

        // Filter only rows that are NOT "Credit Note Uploaded"
        const validRows = updatedRows.filter(
          (row) => row.InvoiceStatus !== "Credit Note Uploaded"
        );

        let runningRemaining = poAmt;

        // Calculate RemainingPoAmount only for valid rows
        updatedRows = updatedRows.map((row) => {
          if (row.InvoiceStatus === "Credit Note Uploaded") {
            // Keep existing RemainingPoAmount as is for credit note rows
            return { ...row };
          }

          const invoiceAmount = parseFloat(row.InvoiceAmount) || 0;

          const updatedRow = {
            ...row,
            RemainingPoAmount: runningRemaining.toFixed(2),
          };

          runningRemaining -= invoiceAmount;
          return updatedRow;
        });

        // Calculate total invoice amount excluding "Credit Note Uploaded"
        const totalInvoiceAmount = validRows.reduce(
          (sum, row) => sum + (parseFloat(row.InvoiceAmount) || 0),
          0
        );

        const remainingAfter = +(poAmt - totalInvoiceAmount).toFixed(2);

        // Find the last valid (non-credit-note) row
        const lastValidRow = [...updatedRows]
          .reverse()
          .find((row) => row.InvoiceStatus !== "Credit Note Uploaded");

        const lastRowHasValue =
          lastValidRow &&
          String(lastValidRow.InvoiceAmount).trim() !== "" &&
          Number(lastValidRow.InvoiceAmount) !== 0;

        // Add a new row if there's remaining amount and the last valid row has a value
        if (poAmt > 0 && remainingAfter > 0 && lastRowHasValue) {
          const maxId =
            updatedRows.length > 0
              ? Math.max(...updatedRows.map((r) => r.id))
              : 0;

          updatedRows.push({
            id: maxId + 1,
            InvoiceDescription: "",
            RemainingPoAmount: remainingAfter.toFixed(2),
            InvoiceAmount: "",
            InvoiceDueDate: "",
            InvoiceProceedDate: "",
            showProceed: false,
            InvoiceStatus: "",
            userInGroup: false,
            employeeEmail: "",
            itemID: null,
            InvoiceNo: "",
            InvoiceDate: "",
            InvoiceTaxAmount: "",
            ClaimNo: null,
            DocId: "",
            InvoiceFileID: "",
            invoiceApprovalChecked: false,
            invoiceCloseApprovalChecked: false, // Initialize here
          });
        }

        // Remove extra rows if total exceeds PO amount
        if (remainingAfter <= 0) {
          updatedRows = updatedRows.filter(
            (row, idx) =>
              idx === 0 ||
              Number(row.InvoiceAmount) !== 0 ||
              idx < updatedRows.length - 1
          );
        }
      }

      // If rowEdit is enabled, update local data
      if (props.rowEdit === "Yes") {
        setLocalInvoiceData(
          updatedRows.map((row) => ({
            InvoiceDescription: row.InvoiceDescription,
            InvoiceAmount: row.InvoiceAmount,
            InvoiceDueDate: row.InvoiceDueDate,
            RemainingPoAmount: row.RemainingPoAmount,
          }))
        );
      }

      console.log("Updated Invoice Rows:", updatedRows);
      return updatedRows;
    });

    console.log(`Field Updated: ${field}, Value: ${value}`);
  };


// Get only selected invoices for closing
const getSelectedCloseInvoices = () =>
  invoiceRows.filter((r) => r.closeInvoiceChecked);

// Validate that sum of selected invoices == closeAmount
const validateCloseAmount = () => {
  const selected = getSelectedCloseInvoices();
  const total = selected.reduce(
    (sum, row) => sum + Number(row.InvoiceAmount || 0),
    0
  );

  if (total !== Number(closeAmount)) {
    alert(
      `Close Amount must be equal to the total Invoice Amounts of selected invoices.\n\nExpected: ${total}`
    );
    return false;
  }

  if (!closeReason.trim()) {
    alert("Close Reason is required.");
    return false;
  }

  return true;
};

const handleDirectCloseInvoices = async () => {
  if (!validateCloseAmount()) return;

  const selected = getPendingCloseInvoices();

  for (const row of selected) {
    if (!row.itemID) continue;

    const requestData = {
      InvoiceStatus: "Closed",
      CloseAmount: closeAmount,
      CloseReason: closeReason,
      CloseDate: moment().format("YYYY-MM-DD"),
      RunWF: "Yes",
    };

    await updateDataToSharePoint(
      InvoiceList,
      requestData,
      props.siteUrl,
      row.itemID
    );
  }

  alert("Invoices closed successfully.");

  // update UI instantly
setInvoiceRows(prev =>
  prev.map(r =>
    selected.some(s => s.itemID === r.itemID)
      ? { ...r, InvoiceStatus: "Closed" }
      : r
  )
);

// refresh from SharePoint
await refreshInvoices();
};
const handleCloseApprove = async () => {
  if (!validateCloseAmount()) return;

  const selected = getSelectedCloseInvoices();

  for (const row of selected) {
    if (!row.itemID) continue;
    const requestData = {
      InvoiceStatus: "Pending Close Approval",
      CloseAmount: closeAmount,
      CloseReason: closeReason,
      PrevInvoiceStatus: row.InvoiceStatus || "Started",
      ManagerDecision: "Pending",
      RunWF: "Yes",
    };


    await updateDataToSharePoint(
      InvoiceList,
      requestData,
      props.siteUrl,
      row.itemID
    );
  }
  alert("Sent for Close Approval.");

  setInvoiceRows(prev =>
  prev.map(r =>
    selected.some(s => s.itemID === r.itemID)
      ? { ...r, InvoiceStatus: "Pending Close Approval" }
      : r
  )
);

await refreshInvoices();

};

const handleUpdateInvoiceRow = async (
  e: React.MouseEvent<HTMLButtonElement>,
  row: InvoiceRow
) => {
  e.preventDefault();

  console.log("=== HANDLE UPDATE INVOICE ROW ===");
  console.log("Row ID:", row.id);
  console.log("Row ItemID:", row.itemID);
  console.log("Row InvoiceStatus:", row.InvoiceStatus);
  console.log("Row employeeEmail:", row.employeeEmail);

  // 🌟 Correct field names from RequestForm
  console.log("SelectedRow EmployeeEmail:", props.selectedRow?.employeeEmail);
  console.log("SelectedRow ManagerEmail:", props.selectedRow?.projectMangerEmail);

  // ------------------------------------------------------------
  // GET TRUE requestor + manager email
  // ------------------------------------------------------------
  const requestorEmail =
    props.selectedRow?.employeeEmail ||
    row.employeeEmail ||
    "";

  const managerEmail =
    props.selectedRow?.projectMangerEmail ||
    ""; // row does NOT contain manager email, only props.selectedRow

  console.log("Resolved Requestor Email:", requestorEmail);
  console.log("Resolved Manager Email:", managerEmail);

  if (!row.itemID) {
    console.error("Missing row.itemID — cannot update SharePoint");
    alert("Error: Invoice row missing item ID.");
    return;
  }

  // -------------------------------------------------------------------
  // CASE 1 — REQUESTOR ≠ MANAGER → send to Pending Manager Approval
  // -------------------------------------------------------------------
  if (requestorEmail && managerEmail && requestorEmail !== managerEmail) {
    console.log("Requester ≠ Manager → Sending to Pending Manager Approval");

    const requestData = {
      InvoiceStatus: "Pending Manager Approval",
      PrevInvoiceStatus: row.InvoiceStatus || "Started",
      ManagerDecision: "Pending",
      RunWF: "Yes",
    };

    console.log("SP Request Payload (Manager Approval):", requestData);

    try {
      const response = await updateDataToSharePoint(
        InvoiceList,
        requestData,
        props.siteUrl,
        row.itemID
      );

      console.log("SP Response (Pending Manager Approval):", response);

      setInvoiceRows((prev) =>
        prev.map((r) =>
          r.id === row.id
            ? { ...r, InvoiceStatus: "Pending Manager Approval" }
            : r
        )
      );

      setProceededRows((prev) => [...prev, row.id]);

      console.log("UI Updated → Row set to Pending Manager Approval");
      alert("Sent for Manager Approval");
    } catch (err) {
      console.error("Error sending for manager approval:", err);
      alert("Failed to send for approval.");
    }

    return; // 🔒 prevents direct proceed path
  }

  // -------------------------------------------------------------------
  // CASE 2 — REQUESTOR = MANAGER → Direct Proceed
  // -------------------------------------------------------------------
  console.log("Requester = Manager → Direct Proceed");

const requestData = {
  ProceedDate:
    row.InvoiceStatus !== "Proceeded"
      ? moment().format("YYYY-MM-DD") // NEW proceed
      : "",                          // keep blank for old unproceeded rows

  InvoiceStatus: "Proceeded",
  PrevInvoiceStatus: row.InvoiceStatus || "Started",
  RunWF: "Yes",
};



  console.log("SP Request Payload (Proceed):", requestData);

  try {
    const response = await updateDataToSharePoint(
      InvoiceList,
      requestData,
      props.siteUrl,
      row.itemID
    );

    console.log("SP Response (Proceeded):", response);

    setInvoiceRows((prev) =>
  prev.map((r) =>
    r.id === row.id
      ? {
          ...r,
          InvoiceStatus: "Proceeded",
          ProceedDate:
            moment(row.InvoiceProceedDate, "DD-MM-YYYY", true).isValid()
              ? moment(row.InvoiceProceedDate, "DD-MM-YYYY").format("DD-MM-YYYY")
              : moment().format("DD-MM-YYYY"),
        }
      : r
  )
);


    setProceededRows((prev) => [...prev, row.id]);

    console.log("UI Updated → Row set to Proceeded");

    alert("Invoice Proceeded Successfully!");
  } catch (error) {
    console.error("ERROR proceeding invoice:", error);
    alert("Failed to proceed invoice.");
  }
};



// // NEW: Requestor proceed → Pending Manager Approval
// const handleRequestorProceedWithManager = async (
//   e: React.MouseEvent<HTMLButtonElement>,
//   row: InvoiceRow
// ) => {
//   e.preventDefault();
//   if (!row.itemID) return;

//   const requestData = {
//     ProceedDate: moment().format("YYYY-MM-DD"),
//     InvoiceStatus: "Pending Manager Approval",
//     ManagerDecision: "",
//     ManagerDecisionDate: null,
//     ManagerReason: "",
//     RunWF: "Yes",
//   };

//   try {
//     await updateDataToSharePoint(InvoiceList, requestData, props.siteUrl, row.itemID);

//     setInvoiceRows((prev) =>
//       prev.map((r) =>
//         r.id === row.id
//           ? { ...r, InvoiceStatus: "Pending Manager Approval" }
//           : r
//       )
//     );

//     alert("Invoice sent to Project Manager for approval.");
//   } catch (err) {
//     console.error("Error updating invoice:", err);
//     alert("Failed to send invoice for manager approval.");
//   }
// };

// NEW: Send Reminder to Project Manager
const handleSendReminder = async (row: InvoiceRow) => {
  if (!row.itemID) return;

  const today = moment().format("YYYY-MM-DD");

  const requestData = {
    ReminderSentDate: today,
    ManagerDecision: "Pending Manager Approval",
    InvoiceStatus: "Pending Manager Approval",
    RunWF: "Yes",
  };

  try {
    await updateDataToSharePoint(
      InvoiceList,
      requestData,
      props.siteUrl,
      row.itemID
    );

    alert("Reminder sent to Project Manager!");

    // update UI
    setInvoiceRows((prev) =>
      prev.map((r) =>
        r.id === row.id
          ? {
              ...r,
              ReminderSentDate: today,
            }
          : r
      )
    );
  } catch (err) {
    console.error("Error sending reminder:", err);
    alert("Failed to send reminder");
  }
};


// NEW: Manager Approves
const handleManagerApprove = async (row: InvoiceRow) => {
  if (!row.itemID) return;

const requestData = {
  InvoiceStatus: "Proceeded",
  ManagerDecision: "Approved",
  ManagerDecisionDate: moment().format("YYYY-MM-DD"),
  ManagerReason: "",
  ProceedDate: moment().format("YYYY-MM-DD"), // always today when manager approves
  RunWF: "Yes",
};



  try {
    await updateDataToSharePoint(InvoiceList, requestData, props.siteUrl, row.itemID);

setInvoiceRows((prev) =>
  prev.map((r) =>
    r.id === row.id
      ? {
          ...r,
          InvoiceStatus: "Proceeded",
          ProceedDate: moment(requestData.ProceedDate).format("DD-MM-YYYY"),
        }
      : r
  )
);



    alert("Invoice approved successfully.");
  } catch (err) {
    console.error("Error:", err);
    alert("Failed to approve invoice.");
  }
};

// NEW: Manager Reject or Hold
const handleManagerDecision = async (
  row: InvoiceRow,
  action: "Reject" | "Hold",
  newDueDate: string,
  reason: string
) => {
  if (!row.itemID) return;

  // -------------------------------
  // BUILD REQUEST DATA BASED ON ACTION
  // -------------------------------
  const requestData: any = {
    ManagerDecision: action,
    ManagerReason: reason,
    ManagerDecisionDate: moment().format("YYYY-MM-DD"),
    RunWF: "Yes",
  };

  if (action === "Reject") {
    // Reject → back to Started + update due date
    requestData.InvoiceStatus = "Started";
    requestData.InvoiceDueDate = newDueDate;
  }

  if (action === "Hold") {
    // Hold → status is On Hold + DO NOT update due date
    requestData.InvoiceStatus = "On Hold";
    // DO NOT add InvoiceDueDate here
  }

  try {
    await updateDataToSharePoint(
      InvoiceList,
      requestData,
      props.siteUrl,
      row.itemID
    );

    // -------------------------------
    // UPDATE UI
    // -------------------------------
    setInvoiceRows((prev) =>
      prev.map((r) =>
        r.id === row.id
          ? {
              ...r,
              InvoiceStatus: requestData.InvoiceStatus,
              // Only update due date for Reject
              ...(action === "Reject" && { InvoiceDueDate: newDueDate }),
            }
          : r
      )
    );

    alert(`Invoice ${action === "Reject" ? "rejected" : "put on hold"}.`);
  } catch (err) {
    console.error("Error:", err);
    alert(`Failed to ${action.toLowerCase()} invoice.`);
  }
};


const handleCloseHoldApprove = async (row: InvoiceRow) => {
  if (!row.itemID) return;

  await updateDataToSharePoint(
    InvoiceList,
    {
      InvoiceStatus: "Closed",
      ManagerCloseDecision: "Approved",
      CloseAmount: row.CloseAmount,
      CloseReason: row.CloseReason,
      CloseDate: moment().format("YYYY-MM-DD"),
      RunWF: "Yes",
    },
    props.siteUrl,
    row.itemID
  );

  setInvoiceRows(prev =>
    prev.map(r =>
      r.id === row.id
        ? { ...r, InvoiceStatus: "Closed" }
        : r
    )
  );

  alert("Invoice closed by manager.");
};



const submitCloseApprove = async () => {
  const rowsToUpdate = getPendingCloseInvoices();

  if (rowsToUpdate.length === 0) {
    alert("No invoices pending close approval");
    return;
  }

  console.debug("submitCloseApprove - rowsToUpdate:", rowsToUpdate);

  for (const row of rowsToUpdate) {
    if (!row.itemID) {
      console.warn("submitCloseApprove - skipping row with no itemID", row);
      continue;
    }
    try {
      const resp = await updateDataToSharePoint(
        InvoiceList,
        {
          InvoiceStatus: "Closed",
          ManagerCloseDecision: "Approved",
          ManagerCloseReason: managerCloseReason || "",
          CloseDate: moment().format("YYYY-MM-DD"),
          RunWF: "Yes",
        },
        props.siteUrl,
        Number(row.itemID)
      );
      console.debug("submitCloseApprove - update resp:", resp);

      // update UI locally so we don't depend on refreshInvoices()
      setInvoiceRows((prev) =>
        prev.map((r) =>
          r.itemID === row.itemID
            ? { ...r, InvoiceStatus: "Closed", ManagerCloseDecision: "Approved", ManagerCloseReason: managerCloseReason || "", CloseDate: moment().format("YYYY-MM-DD") }
            : r
        )
      );
    } catch (err) {
      console.error("submitCloseApprove - failed to update row", row, err);
      alert("Failed to close one or more invoices. See console for details.");
    }
  }

  alert("Invoices Closed Successfully");

  // only refresh if we have a RequestID; otherwise skip to avoid the Missing RequestID error
  if (props.selectedRow?.RequestID) {
    await refreshInvoices();
  }
};

const submitCloseReject = async () => {
  if (!managerCloseReason.trim()) {
    alert("Reason is required");
    return;
  }

  if (!modalRow || !modalRow.itemID) return;

  const newStatus = modalRow.PrevInvoiceStatus || "Started";

  await updateDataToSharePoint(
    InvoiceList,
    {
      InvoiceStatus: newStatus,
      ManagerCloseDecision: "Rejected",
      ManagerCloseReason: managerCloseReason,
      RunWF: "Yes",
    },
    props.siteUrl,
    modalRow.itemID
  );

  setInvoiceRows(prev =>
    prev.map(r =>
      r.itemID === modalRow.itemID
        ? { ...r, InvoiceStatus: newStatus }
        : r
    )
  );

  alert("Close Request Rejected");
  setManagerCloseReason("");
  setShowCloseRejectModal(false);

  await refreshInvoices();
};



const submitCloseHold = async () => {
  if (!managerCloseReason.trim()) {
    alert("Reason is required");
    return;
  }

  const pending = getPendingCloseInvoices();
  console.debug("submitCloseHold - pending:", pending);

  for (const row of pending) {
    if (!row.itemID) {
      console.warn("submitCloseHold - skipping row with no itemID", row);
      continue;
    }

    try {
      const resp = await updateDataToSharePoint(
        InvoiceList,
        {
          InvoiceStatus: "Close Hold",
          ManagerCloseDecision: "Hold",
          ManagerCloseReason: managerCloseReason,
          RunWF: "Yes",
        },
        props.siteUrl,
        Number(row.itemID)
      );
      console.debug("submitCloseHold - update resp:", resp);

      // update UI locally
      setInvoiceRows((prev) =>
        prev.map((r) =>
          r.itemID === row.itemID
            ? { ...r, InvoiceStatus: "Close Hold", ManagerCloseDecision: "Hold", ManagerCloseReason: managerCloseReason }
            : r
        )
      );
    } catch (err) {
      console.error("submitCloseHold - failed to update row", row, err);
      alert("Failed to put one or more invoices on hold. See console for details.");
    }
  }

  alert("Close Request Put On Hold");
  setManagerCloseReason("");
  setShowCloseHoldModal(false);

  if (props.selectedRow?.RequestID) {
    await refreshInvoices();
  }
};




const submitCloseHoldReject = async () => {
  if (!managerCloseReason.trim()) {
    alert("Reason required");
    return;
  }

  if (!modalRow || !modalRow.itemID) return;

  await updateDataToSharePoint(
    InvoiceList,
    {
      InvoiceStatus: modalRow.PrevInvoiceStatus || "Started",
      ManagerCloseDecision: "Rejected",
      ManagerCloseReason: managerCloseReason,
      RunWF: "Yes",
    },
    props.siteUrl,
    modalRow.itemID
  );

  alert("Close Request Rejected");
  setShowCloseRejectModal(false);
  setManagerCloseReason("");

  await refreshInvoices();
};


const showCloseInvoiceColumn =
  invoiceRows.some(
    (row) =>
      row.InvoiceStatus === "Started" ||
      row.InvoiceStatus === "Proceeded" ||
      row.InvoiceStatus === "Close Hold" ||
      row.InvoiceStatus === "Pending Close Approval"
  );

const showCloseInvoicesSection =
  invoiceRows.some(
    (row) =>
      row.InvoiceStatus === "Started" ||
      row.InvoiceStatus === "Proceeded" ||
      row.InvoiceStatus === "Close Hold" ||
      row.InvoiceStatus === "Pending Close Approval"
  );

// When user manually selects checkboxes, auto-update Close Amount
// const handleCheckboxToggle = (itemID: number) => {
//   setInvoiceRows(prev =>
//     prev.map(r =>
//       r.itemID === itemID
//         ? { ...r, closeInvoiceChecked: !r.closeInvoiceChecked }
//         : r
//     )
//   );

//   // After updating selection, recalc the close amount
//   setTimeout(() => {
//     const selected = invoiceRows
//       .filter(r => r.closeInvoiceChecked || r.itemID === itemID) // Include toggled one
//       .map(r => Number(r.InvoiceAmount) || 0);

//     const newTotal = selected.reduce((a, b) => a + b, 0);

//     setCloseAmount(String(newTotal));
//   }, 0);
// };




  // console.log(invoiceRows, "invoiceRowsabc"); // Log invoiceRows to check its value
  // console.log(5, "approverStatusinvoiceRowsabc"); // Log invoiceRows to check its value
  const handleHistoryClick = async (
    e: React.MouseEvent<HTMLButtonElement>,
    row: any
  ) => {
    e.preventDefault();
    // console.log(`History clicked for row ${row.itemID}`);
    // const filterQuery = `$filter=CMSRequestItemID eq '${row.itemID}'&$orderby=Id desc`;
    const filterQuery = `$select=*,Author/Title&$expand=Author&$filter=CMSRequestItemID eq '${row.itemID}'&$orderby=Id desc`;
    setSelectedRow(row); // Set the selected row to access its itemID

    setHistoryLoading(true);
    setShowHistoryModal(true); // Show modal before fetching (or after if you prefer)

    try {
      const response = await getSharePointData(
        { context },
        InvoiceHistory,
        filterQuery
      );
      // console.log("Invoice history fetched successfully:", response);
      setInvoiceHistoryData(response); // Store history data
      // console.log(invoiceHistoryData, "invoiceHistoryData");
      // console.log()
    } catch (error) {
      console.error("Error fetching invoice history:", error);
      setInvoiceHistoryData([]);
    } finally {
      setHistoryLoading(false);
    }
  };

  const handleCloseModal = () => {
    setShowEditModal(false); // Close the modal
    setSelectedRow(null); // Clear the selected row data
  };

  const proceedButtonCount = invoiceRows.filter(
    (row) => row.InvoiceStatus === "Started"
  ).length;
  React.useEffect(() => {
    // console.log("proceedButtonCount:", proceedButtonCount);
    if (onProceedButtonCountChange) {
      onProceedButtonCountChange(proceedButtonCount);
    }
  }, [proceedButtonCount, onProceedButtonCountChange]);

  // (row add/remove handled in parent `RequestForm`)
  React.useEffect(() => {
    setLocalInvoiceData(
      invoiceRows.map((row) => ({
        InvoiceDescription: row.InvoiceDescription,
        InvoiceAmount: row.InvoiceAmount,
        InvoiceDueDate: row.InvoiceDueDate,
        RemainingPoAmount: row.RemainingPoAmount,
      }))
    );
  }, [invoiceRows]);

  // Function to handle local state updates for all fields
  const handleLocalFieldChange = (
    index: number,
    field: keyof InvoiceRow,
    value: string | number
  ) => {
    setLocalInvoiceData((prev) => {
      const updatedData = [...prev];
      updatedData[index] = { ...updatedData[index], [field]: value };
      return updatedData;
    });

    // Call handleInvoiceChange to update parent state
    handleInvoiceChange(index, field, value);
    // if (field === "InvoiceAmount") {
    //   handleInvoiceTextChange(index, field, value);
    // }
  };

  // Initialize local state for all fields
  React.useEffect(
    () => {
      // Ensure we only initialize once when entering row-edit + Approved state.
      // Calling handleTextFieldChange here caused updates to invoiceRows which re-triggered this effect.
      // Instead, only initialize local state (setLocalInvoiceData) and avoid mutating invoiceRows.
      // const initializedRef = (React as any).useRef?.current; // placeholder to satisfy TS transpile below
    },
    [
      /* intentionally left blank to be replaced by the block below */
    ]
  );

  // Replacement effect: run once when rowEdit becomes "Yes" and approverStatus is "Approved".
  const _initForApprovedRef = React.useRef(false);
  React.useEffect(() => {
    if (
      props.rowEdit === "Yes" &&
      props.selectedRow?.approverStatus === "Approved"
    ) {
      if (_initForApprovedRef.current) return; // already initialized

      // Build local data from invoiceRows without calling handlers that mutate invoiceRows
      const initialLocal = invoiceRows.map((row) => ({
        InvoiceDescription: row.InvoiceDescription,
        InvoiceAmount: row.InvoiceAmount,
        InvoiceDueDate: row.InvoiceDueDate,
        RemainingPoAmount: row.RemainingPoAmount,
      }));

      setLocalInvoiceData(initialLocal);

      // Mark as initialized so we don't re-run and cause a loop
      _initForApprovedRef.current = true;
    } else {
      // Reset so effect can run again when conditions meet in the future
      _initForApprovedRef.current = false;
    }
    // NOTE: we intentionally depend only on props.rowEdit and approverStatus so this runs
    // when the approval state changes. We avoid depending on invoiceRows to prevent loops.
  }, [props.rowEdit, props.selectedRow?.approverStatus]);
  const hasCalledLocalFieldChangeRef = React.useRef(false);

  // React.useEffect(() => {
  //   if (props.rowEdit === "Yes" && !hasCalledLocalFieldChangeRef.current) {
  //     const timeoutId = setTimeout(() => {
  //       invoiceRows.forEach((row, index) => {
  //         if (row.CreditNoteStatus === "Uploaded") {
  //           // Call handleLocalFieldChange for each row with CreditNoteStatus "Uploaded"
  //           handleLocalFieldChange(index, "InvoiceAmount", row.InvoiceAmount);
  //           handleLocalFieldChange(index, "InvoiceDescription", row.InvoiceDescription);
  //           handleLocalFieldChange(index, "InvoiceDueDate", row.InvoiceDueDate);
  //           handleLocalFieldChange(index, "RemainingPoAmount", row.RemainingPoAmount);
  //         }
  //       });
  //       hasCalledLocalFieldChangeRef.current = true; // Mark as called
  //     }, 1000);
  //     //kojioh
  //     return () => clearTimeout(timeoutId); // Cleanup timeout on unmount or re-run
  //   }
  // }, [props.rowEdit, invoiceRows]);
  const syncUploadedCreditNoteRows = () => {
    invoiceRows.forEach((row, index) => {
      // if (row.CreditNoteStatus === "Uploaded") {
      handleLocalFieldChange(index, "InvoiceAmount", row.InvoiceAmount);
      // handleLocalFieldChange(
      //   index,
      //   "InvoiceDescription",
      //   row.InvoiceDescription
      // );
      // handleLocalFieldChange(index, "InvoiceDueDate", row.InvoiceDueDate);
      handleLocalFieldChange(index, "RemainingPoAmount", row.RemainingPoAmount);
      // }
    });
    hasCalledLocalFieldChangeRef.current = true;
  };
  React.useEffect(() => {
    if (props.rowEdit === "Yes" && !hasCalledLocalFieldChangeRef.current) {
      const timeoutId = setTimeout(() => {
        syncUploadedCreditNoteRows();
      }, 1000);
      return () => clearTimeout(timeoutId);
    }
  }, [props.rowEdit, invoiceRows]);


// ======================================================
// AUTO SELECT CLOSE INVOICES BASED ON CloseAmount
// ======================================================
React.useEffect(() => {
  if (!closeAmount || Number(closeAmount) <= 0) {
    // If cleared → unselect all
    setInvoiceRows(prev =>
      prev.map(r => ({ ...r, closeInvoiceChecked: false }))
    );
    return;
  }

  const target = Number(closeAmount);

  // Sort rows bottom → top BUT ONLY Started + Proceeded allowed
  const sorted = [...invoiceRows]
    .filter(r =>
      ["Started", "Proceeded"].includes(r.InvoiceStatus)
    )
    .sort((a, b) => b.id - a.id);

  let running = 0;
  const rowsToSelect = new Set<number>();

  for (const row of sorted) {
    const amt = Number(row.InvoiceAmount || 0);

    if (running + amt <= target) {
      running += amt;
      rowsToSelect.add(row.id);
    }

    if (running >= target) break;
  }

  setInvoiceRows(prev =>
    prev.map(r =>
      rowsToSelect.has(r.id)
        ? { ...r, closeInvoiceChecked: true }
        : { ...r, closeInvoiceChecked: false }
    )
  );
}, [closeAmount]);


  return (
    <div className="mt-4">
      <div
        className="d-flex justify-content-between align-items-center mb-3 sectionheader"
        style={{ padding: "7px 8px" }}
      >
        <div className="d-flex align-items-center justify-content-between">
          {/* Invoice section approval checkbox (editable mode shown by parent) */}
          {isEditMode &&
            props.selectedRow &&
            props.selectedRow.employeeEmail === currentUserEmail &&
            props.selectedRow.isPaymentReceived !== "Yes" &&
            !["Approved", "Hold", "Pending From Approver", "Reminder"].includes(
              props.selectedRow.approverStatus
            ) &&
            props.selectedRow.isCreditNoteUploaded !== "No" && (

              <span
                className="form-check me-2"
                style={{
                  display: "flex",
                  alignItems: "center",
                  marginRight: 8,
                }}
              >
                <input
                  type="checkbox"
                  id="cbInvoice"
                  className="form-check-input"
                  checked={invoiceApprovalChecked}
                  onChange={(e) => {
                    setInvoiceApprovalChecked &&
                      setInvoiceApprovalChecked(e.target.checked);
                    setShowEditInvoiceColumn(e.target.checked); // Toggle visibility
                  }}
                  onClick={(ev) => ev.stopPropagation()}
                />
              </span>
            )}

          <h5
            className="fw-bold mt-2 me-2 headingColor"
            style={{ cursor: "pointer" }}
            onClick={() => setIsCollapsed((prev) => !prev)}
            aria-expanded={isCollapsed}
            aria-controls="poDetailsCollapse"
          >
            Invoice Details
          </h5>
        </div>
      </div>

      {/* Responsive Table */}
      <div
        className={`${
          isCollapsed ? "collapse show" : "collapse"
        } sectioncontent`}
        id="poDetailsCollapse"
      >
        <style>{`
          .tablescrollwrapper {
            overflow-x: auto;
            width: 100%;
          }
          .fixedcolumn, .fixed-th {
            min-width: 180px;
            max-width: 240px;
            width: 200px;
            white-space: nowrap;
          }
          .fixed-serial {
            min-width: 80px;
            max-width: 100px;
            width: 90px;
            white-space: nowrap;
          }
        `}</style>
        <div className="tablescrollwrapper">
          <table
            className="table table-bordered align-middle"
            style={{ minWidth: "1200px" }}
          >
            <thead className="table-light">
              <tr>
                <th className="fixed-th fixed-serial">S.No</th>
                {/* <th className="fixed-th "><input
                  type="checkbox"
                  style={{ marginLeft: 8 }}
                  checked={allSelected}
                  onChange={handleSelectAll}
                /></th> */}

                <th className="fixed-th">
                  Invoice Description <span style={{ color: "red" }}>*</span>
                </th>
                <th className="fixed-th">Remaining PO Amount</th>
                <th className="fixed-th">
                  Invoice Amount <span style={{ color: "red" }}>*</span>
                </th>
                <th className="fixed-th">
                  Invoice Due Date <span style={{ color: "red" }}>*</span>
                </th>
                {invoiceRows.some((row) => row.showProceed) && (
                  <th className="fixed-th">Invoice Proceed Date</th>
                )}
                {invoiceRows.some(
                  (row) =>
                    row.InvoiceStatus === "Generated" ||
                    row.InvoiceStatus === "Credit Note Uploaded" ||
                    row.PrevInvoiceStatus === "Generated"
                ) && <th className="">Invoice Attachment</th>}
                {/* <th className="">Invoice Status</th> */}
                {props.rowEdit === "Yes" && (
                  <th className="">Invoice Status</th>
                )}
                <th className="fixed-th">Action</th>

                {showCloseInvoiceColumn && (
                <th className="fixed-th">
                  <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                    <input
                      type="checkbox"
                      className="form-check-input"
                      checked={
                        invoiceRows
                          .filter(
                            (r) =>
                              r.InvoiceStatus === "Started" ||
                              r.InvoiceStatus === "Proceeded"
                          )
                          .every((r) => r.closeInvoiceChecked)
                      }
                        onChange={(e) => {
                          const checked = e.target.checked;

                          // 1. Update all eligible rows
                          setInvoiceRows((prevRows) =>
                            prevRows.map((r) =>
                              r.InvoiceStatus === "Started" || r.InvoiceStatus === "Proceeded"
                                ? { ...r, closeInvoiceChecked: checked }
                                : r
                            )
                          );

                          // 2. Recalculate Close Amount automatically
                          setTimeout(() => {
                            const selectableRows = invoiceRows.filter(
                              (r) =>
                                r.InvoiceStatus === "Started" || r.InvoiceStatus === "Proceeded"
                            );

                            const total = checked
                              ? selectableRows.reduce(
                                (sum, r) => sum + (Number(r.InvoiceAmount) || 0),
                                0
                              )
                              : 0;

                            setCloseAmount(String(total));
                          }, 0);
                        }}

                    />

                    <span>Close Invoice</span>
                  </div>
                </th>
                )}

                {/* Add a new column header "Edit Invoice" if the checkbox condition is met */}
                {showEditInvoiceColumn && (
                  <th className="fixed-th">
                    <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                      <input
                        type="checkbox"
                        className="form-check-input"
                        checked={invoiceRows
                          .filter(r =>
                            !(
                              r.InvoiceStatus === "Credit Note Uploaded" ||
                              r.InvoiceStatus === "Pending Approval" ||
                              r.CreditNoteStatus === "Pending"
                            )
                          )
                          .every(r => r.invoiceApprovalChecked)
                        }
                        onChange={(e) => {
                          const checked = e.target.checked;

                          setInvoiceRows((prevRows) =>
                            prevRows.map((r) =>
                              !(
                                r.InvoiceStatus === "Credit Note Uploaded" ||
                                r.InvoiceStatus === "Pending Approval" ||
                                r.CreditNoteStatus === "Pending"
                              )
                                ? { ...r, invoiceApprovalChecked: checked }
                                : r
                            )
                          );
                        }}
                      />
                      <span>Edit Invoice</span>
                    </div>
                  </th>
                )}
              </tr>
            </thead>
            <tbody>
              {invoiceRows
                .filter((row) => {
                  // Debugging log to verify rows being rendered
                  console.log("Rendering row:", row);

                  // If user belongs to CMSAccountGroup → apply filter
                  if (userGroups.includes("CMSAccountGroup")) {
                    const invoiceVisible = ![
                      "Started",
                      "Pending Approval",
                      "",
                    ].includes(row.InvoiceStatus);

                    const approver =
                      props.selectedRow?.approverStatus || approverStatus;
                    const creditNotePendingVisible =
                      (approver === "Approved" || approver === "Completed") &&
                      row.CreditNoteStatus === "Pending";

                    return invoiceVisible || creditNotePendingVisible;
                  }

                  // Otherwise → show all rows
                  return true;
                })
                .slice()
                .sort((a, b) => {
                  const claimA =
                    a.ClaimNo !== null
                      ? Number(a.ClaimNo)
                      : Number.MAX_SAFE_INTEGER;
                  const claimB =
                    b.ClaimNo !== null
                      ? Number(b.ClaimNo)
                      : Number.MAX_SAFE_INTEGER;

                  // If both ClaimNo are null, sort by id
                  if (
                    claimA === Number.MAX_SAFE_INTEGER &&
                    claimB === Number.MAX_SAFE_INTEGER
                  ) {
                    return a.id - b.id;
                  }

                  return claimA - claimB;
                }) // .map((row, index) => (
                .map((row, index) => {
                  const managerEmail = props.selectedRow?.projectMangerEmail;
                  return (
                  <tr key={row.id}>
                    <td className="fixedcolumn fixed-serial">{index + 1}</td>
                    {/* <td className="fixedcolumn "> <input
            type="checkbox"
            style={{ marginLeft: 8 }}
            checked={selectedRows.includes(row.id)}
            onChange={handleSelectRow(row.id)}
          /></td> */}
                    <td className="fixedcolumn">
                      <textarea
                        className={`form-control ${
                          errors[`InvoiceDescription_${index}`]
                            ? "is-invalid"
                            : ""
                        }`}
                        value={
                          props.rowEdit === "Yes"
                            ? localInvoiceData[index]?.InvoiceDescription || "" // Keep it blank if cleared
                            : row.InvoiceDescription
                        }
                        onChange={(e) => {
                          const value = e.target.value;
                          // if (props.rowEdit === "Yes") {
                          //   handleLocalFieldChange(
                          //     index,
                          //     "InvoiceDescription",
                          //     value
                          //   );
                          // }

                          if (props.rowEdit === "Yes") {
                            handleTextFieldChange(
                              index,
                              "InvoiceDescription",
                              value
                            );

                            handleLocalFieldChange(
                              index,
                              "InvoiceDescription",
                              value
                            );
                          } else {
                            handleTextFieldChange(
                              index,
                              "InvoiceDescription",
                              value
                            );
                          }
                        }}
                        disabled={
                          props.rowEdit === "Yes"
                            ? !(
                                props.selectedRow?.employeeEmail ===
                                  currentUserEmail &&
                                props.selectedRow?.selectedSections
                                  ?.toLowerCase()
                                  .includes("invoice") &&
                                props.selectedRow?.approverStatus ===
                                  "Approved" &&
                                (row.InvoiceStatus === "" ||
                                  (row.InvoiceStatus === "Pending Approval" &&
                                    row.CreditNoteStatus === ""))
                              )
                            : false
                        }
                      />
                      {errors[`InvoiceDescription_${index}`] && (
                        <div className="invalid-feedback">
                          {errors[`InvoiceDescription_${index}`]}
                        </div>
                      )}
                    </td>
                    <td className="fixedcolumn">
                      <input
                        type="text"
                        className="form-control"
                        // value={
                        //   index === 0
                        //     ? totalPoAmount.toFixed(2)
                        //     : row.RemainingPoAmount
                        // }
                        value={
                          index === 0
                            ? totalPoAmount.toFixed(2) // Always set totalPoAmount for the first row
                            : props.rowEdit === "Yes"
                            ? localInvoiceData[index]?.RemainingPoAmount || "" // Use local data if in rowEdit mode
                            : row.RemainingPoAmount // Use the calculated RemainingPoAmount for other rows
                        }
                        disabled
                      />
                    </td>
                    <td className="fixedcolumn">
                      <input
                        type="number"
                        className={`form-control ${
                          errors[`InvoiceAmount_${index}`] ? "is-invalid" : ""
                        }`}
                        value={
                          props.rowEdit === "Yes"
                            ? localInvoiceData[index]?.InvoiceAmount || "" // Keep it blank if cleared
                            : row.InvoiceAmount
                        }
                        min={0}
                        step="any"
                        onChange={(e) => {
                          const value = e.target.value;
                          if (props.rowEdit === "Yes") {
                            handleTextFieldChange(
                              index,
                              "InvoiceAmount",
                              value
                            );

                            handleLocalFieldChange(
                              index,
                              "InvoiceAmount",
                              value
                            );
                          } else {
                            handleTextFieldChange(
                              index,
                              "InvoiceAmount",
                              value
                            );
                          }
                        }}
                        disabled={
                          props.rowEdit === "Yes"
                            ? !(
                                props.selectedRow?.employeeEmail ===
                                  currentUserEmail &&
                                props.selectedRow?.selectedSections
                                  ?.toLowerCase()
                                  .includes("invoice") &&
                                props.selectedRow?.approverStatus ===
                                  "Approved" &&
                                (row.InvoiceStatus === "" ||
                                  (row.InvoiceStatus === "Pending Approval" &&
                                    row.CreditNoteStatus === ""))
                              )
                            : false
                        }
                      />
                      {errors[`InvoiceAmount_${index}`] && (
                        <div className="invalid-feedback">
                          {errors[`InvoiceAmount_${index}`]}
                        </div>
                      )}
                    </td>
                    <td className="fixedcolumn">
                      <DatePicker
                        format="DD-MM-YYYY"
                        value={
                          props.rowEdit === "Yes"
                            ? localInvoiceData[index]?.InvoiceDueDate
                              ? moment(
                                  localInvoiceData[index]?.InvoiceDueDate,
                                  "DD-MM-YYYY"
                                )
                              : null // Keep it blank if cleared
                            : row.InvoiceDueDate
                            ? moment(row.InvoiceDueDate, "DD-MM-YYYY")
                            : null
                        }
                        onChange={(date) => {
                          const value = date ? date.format("DD-MM-YYYY") : ""; // Handle null value
                          // if (props.rowEdit === "Yes") {
                          //   handleLocalFieldChange(
                          //     index,
                          //     "InvoiceDueDate",
                          //     value
                          //   );
                          // }

                          if (props.rowEdit === "Yes") {
                            handleTextFieldChange(
                              index,
                              "InvoiceDueDate",
                              value
                            );

                            handleLocalFieldChange(
                              index,
                              "InvoiceDueDate",
                              value
                            );
                          } else {
                            handleTextFieldChange(
                              index,
                              "InvoiceDueDate",
                              value
                            );
                          }
                        }}
                        disabledDate={(current) =>
                          current && current < moment().startOf("day")
                        }
                        disabled={
                          props.rowEdit === "Yes"
                            ? !(
                                props.selectedRow?.employeeEmail ===
                                  currentUserEmail &&
                                props.selectedRow?.selectedSections
                                  ?.toLowerCase()
                                  .includes("invoice") &&
                                props.selectedRow?.approverStatus ===
                                  "Approved" &&
                                (row.InvoiceStatus === "" ||
                                  (row.InvoiceStatus === "Pending Approval" &&
                                    row.CreditNoteStatus === ""))
                              )
                            : false
                        }
                      />
                      {errors[`InvoiceDueDate_${index}`] && (
                        <div className="invalid-feedback">
                          {errors[`InvoiceDueDate_${index}`]}
                        </div>
                      )}
                    </td>
                    {/* {row.showProceed && (
                      <td className="fixedcolumn">
                        <DatePicker
                          type="date"
                          format="DD-MM-YYYY"
                          value={
                            !row.InvoiceProceedDate ||
                            row.InvoiceProceedDate === "01/01/1970"
                              ? moment()
                              : moment(row.InvoiceProceedDate, "DD-MM-YYYY")
                          }
                          onChange={(date) =>
                            handleTextFieldChange(
                              index,
                              "InvoiceProceedDate",
                              date ? date.format("DD-MM-YYYY") : ""
                            )
                          }
                          className="form-control"
                          disabled={isEditMode}
                        />
                      </td>
                      )} */}
                    {invoiceRows.some((r) => r.showProceed) && (
                      <td className="fixedcolumn">
                        {row.showProceed ? (
                          <DatePicker
                            format="DD-MM-YYYY"
                            value={
                              !row.InvoiceProceedDate ||
                              row.InvoiceProceedDate === "01/01/1970"
                                ? moment()
                                : moment(row.InvoiceProceedDate, "DD-MM-YYYY")
                            }
                            onChange={(date) =>
                              handleTextFieldChange(
                                index,
                                "InvoiceProceedDate",
                                date ? date.format("DD-MM-YYYY") : ""
                              )
                            }
                            className="form-control"
                            disabled={isEditMode}
                          />
                        ) : // empty cell when this row doesn't have proceed date
                        null}
                      </td>
                    )}
                    {invoiceRows.some(
                      (r) =>
                        r.InvoiceStatus === "Generated" ||
                        r.InvoiceStatus === "Credit Note Uploaded" ||
                        r.PrevInvoiceStatus === "Generated"
                    ) && (
                      <td className="">
                        {row.InvoiceStatus === "Generated" ||
                        row.InvoiceStatus === "Credit Note Uploaded" ||
                        row.PrevInvoiceStatus === "Generated" ? (
                          row.InvoiceFileID ? (
                            (() => {
                              const file = invoiceDocuments.find(
                                (doc) => doc.DocID === row.InvoiceFileID
                              );
                              return file ? (
                                <button
                                  type="button"
                                  className="btn btn-link"
                                  onClick={(e) =>
                                    handleDownload(e, file.EncodedAbsUrl, {
                                      context,
                                    })
                                  }
                                >
                                  {file.FileLeafRef}
                                </button>
                              ) : (
                                <span>No file found</span>
                              );
                            })()
                          ) : (
                            <span>No file found</span>
                          )
                        ) : (
                          <span>Invoice Not Generated</span>
                        )}
                      </td>
                    )}
                    {/* {props.rowEdit === "Yes" && (
                    <td> {row.InvoiceStatus || "-"}</td>
                    )} */}
                    {props.rowEdit === "Yes" && (
                      <td>
                        {(() => {
                          // If both pending approval and credit note pending, show explicit message
                          if (
                            row.InvoiceStatus === "Pending Approval" &&
                            row.CreditNoteStatus === "Pending"
                          ) {
                            return (
                              <span
                                className="badge rounded-pill px-3 py-2 text-capitalize bg-danger text-white"
                                style={{
                                  display: "inline-block",
                                  textAlign: "center",
                                }}
                              >
                                Credit Note Not Uploaded
                              </span>
                            );
                          }

                          // Determine "On Editing" state for Approved / Completed approver statuses
                          const isOnEditing =
                            (props.selectedRow?.approverStatus === "Approved" &&
                              (row.InvoiceStatus === "Pending Approval" ||
                                row.CreditNoteStatus === "Pending")) ||
                            (props.selectedRow?.approverStatus ===
                              "Completed" &&
                              (row.InvoiceStatus === "Pending Approval" ||
                                row.CreditNoteStatus === "Pending"));

                          const displayStatus = isOnEditing
                            ? "On Editing"
                            : row.InvoiceStatus || "Started";

                          const getBadgeClass = (status: string) => {
                            switch (status) {
                              case "Started":
                                return "bg-primary text-white";
                              case "Proceeded":
                                return "bg-warning text-dark";
                              case "Generated":
                                return "bg-info text-dark";
                              case "Credit Note Uploaded":
                                return "bg-success text-white";
                              case "Invoice Closed":
                                return "bg-danger text-white";
                              case "On Editing":
                                return "bg-warning text-dark";
                              case "Credit Note Not Uploaded":
                                return "bg-danger text-white";
                              default:
                                return "bg-secondary text-white";
                            }
                          };

                          return (
                            <span
                              className={`badge rounded-pill px-3 py-2 text-capitalize ${getBadgeClass(
                                displayStatus
                              )}`}
                              style={{
                                display: "inline-block",
                                textAlign: "center",
                              }}
                            >
                              {displayStatus}
                            </span>
                          );
                        })()}
                      </td>
                    )}

                    <td className="fixedcolumn">
                        {/* {(() => {
                          console.log("Render Check for Proceed:", {
                            isStarted: row.InvoiceStatus === "Started",
                            notProceeded: !proceededRows.includes(row.id),
                            emailMatch: currentUserEmail === row.employeeEmail,
                            showProceed: row.showProceed,
                            fullCondition:
                              row.InvoiceStatus === "Started" &&
                              !proceededRows.includes(row.id) &&
                              currentUserEmail === row.employeeEmail,
                          });
                          return null;
                        })()}
                      */}
                      {isEditMode &&  (
                        <>
                          {/* {row.InvoiceStatus === "Started" &&
                            !proceededRows.includes(row.id) &&
                            row.employeeEmail === currentUserEmail && (
                              // !pendingStatuses.includes(approverStatus) &&
                              <button
                                className="btn btn-primary me-2"
                                onClick={(e) => handleUpdateInvoiceRow(e, row)}
                              >
                                Proceed
                              </button>
                            )} */}

                            
                            {row.InvoiceStatus === "Started" &&
                              !proceededRows.includes(row.id) &&
                              currentUserEmail === row.employeeEmail && (
                                <button
                                  className="btn btn-primary me-2"
                                  onClick={(e) => handleUpdateInvoiceRow(e, row)}
                                >
                                  Proceed
                                </button>
                              )}
                            {/* === MANAGER ACTION BUTTONS (PATCH 2) === */}
                            {managerEmail &&
                              currentUserEmail === managerEmail &&
                              ["Pending Manager Approval", "On Hold"].includes(row.InvoiceStatus) && (
                              <div className="d-flex mb-2">
                                <button
                                  className="btn btn-success btn-sm me-2"
                                  onClick={() => handleManagerApprove(row)}
                                >
                                  Approve
                                </button>

                                <button
                                  className="btn btn-danger me-sm me-2"
                                  onClick={() => openRejectModal(row)} // will be added in phase 3
                                >
                                  Reject
                                </button>

                                <button
                                  className="btn btn-warning btn-sm"
                                  onClick={() => openHoldModal(row)} // phase 3
                                >
                                  Hold
                                </button>
                              </div>
                            )}

                            {/* ================= CLOSE HOLD → MANAGER ACTION BUTTONS IN ACTION COLUMN =============== */}
{currentUserEmail === managerEmail && row.InvoiceStatus === "Close Hold" && (
  <div className="d-flex mb-2">

    <button
  className="btn btn-success btn-sm me-2"
  onClick={() => handleCloseHoldApprove(row)}
>
  Approve
</button>

<button
  className="btn btn-danger btn-sm me-2"
  onClick={() => {
    setModalRow(row);
    setShowCloseRejectModal(true);
  }}
>
  Reject
</button>

<button
  className="btn btn-warning btn-sm"
  onClick={() => {
    setModalRow(row);
    setShowCloseHoldModal(true);
  }}
>
  Hold
</button>


  </div>
)}


                            {row.InvoiceStatus === "Pending Manager Approval" &&
                            row.employeeEmail === currentUserEmail && (
                              <button
                                className="btn btn-warning btn-sm"
                                style={{ marginTop: "4px" }}
                                onClick={() => handleSendReminder(row)}
                              >
                                Send Reminder
                              </button>
                            )}
                          <button
                            className="btn btn-secondary me-2"
                            onClick={(e) => handleHistoryClick(e, row)}
                          >
                            <FontAwesomeIcon
                              icon={faClockRotateLeft}
                              title="Invoice History"
                            />
                          </button>
                        </>
                      )}
                      {isEditMode &&
                        row.employeeEmail === currentUserEmail &&
                        row.InvoiceStatus === "Proceed Approval" && (
                          <button className="btn btn-success" type="button">
                            Proceed approval pending
                          </button>
                        )}
                      {/* {!isEditMode && ( */}
                      {/* {(!isEditMode ||
                        (props.rowEdit === "Yes" &&
                          props.selectedRow?.employeeEmail ===
                            currentUserEmail &&
                          props.selectedRow?.selectedSections
                            ?.toLowerCase()
                            .includes("invoice") &&
                          props.selectedRow?.approverStatus === "Approved" &&
                          row.InvoiceStatus === "Pending Approval" &&
                          row.PrevInvoiceStatus !== "Generated")) && (
                        <button
                          className="btn btn-danger"
                          onClick={() => deleteInvoiceRow(row.id)}
                          disabled={invoiceRows.length === 1}
                          title="Delete Invoice Row"
                        >
                          <FontAwesomeIcon icon={faTrash} />
                        </button>
                      )} */}
                      {(!isEditMode ||
                        (props.rowEdit === "Yes" &&
                          props.selectedRow?.employeeEmail ===
                            currentUserEmail &&
                          props.selectedRow?.selectedSections
                            ?.toLowerCase()
                            .includes("invoice") &&
                          props.selectedRow?.approverStatus === "Approved" &&
                          row.InvoiceStatus === "Pending Approval" &&
                          row.PrevInvoiceStatus !== "Generated" &&
                          // ensure CreditNoteStatus has no value
                          (!row.CreditNoteStatus ||
                            row.CreditNoteStatus === ""))) && (
                        <button
                          className="btn btn-danger"
                          // onClick={() => deleteInvoiceRow(row.id)}
                          onClick={(e) => {
                            e.preventDefault();
                            deleteInvoiceRow(row.id);
                          }}
                          disabled={invoiceRows.length === 1}
                          title="Delete Invoice Row"
                        >
                          <FontAwesomeIcon icon={faTrash} />
                        </button>
                      )}
                    </td>
                    {showCloseInvoiceColumn && (
                      <td className="fixedcolumn">
                        {/* Requestor can select Started/Proceeded invoices */}
                        {currentUserEmail === requestorEmail &&
                          (row.InvoiceStatus === "Started" || row.InvoiceStatus === "Proceeded") && (
                            <input
                              type="checkbox"
                              className="form-check-input"
                              checked={row.closeInvoiceChecked || false}
                              onChange={(e) => {
                                const checked = e.target.checked;

                                // 1. KEEP OLD FUNCTIONALITY (no change)
                                setInvoiceRows((prevRows) =>
                                  prevRows.map((r) =>
                                    r.id === row.id
                                      ? { ...r, closeInvoiceChecked: checked }
                                      : r
                                  )
                                );

                                // 2. NEW FUNCTIONALITY — auto-update Close Amount
                                setTimeout(() => {
                                  const selectedRows = invoiceRows
                                    .filter((r) => r.closeInvoiceChecked || r.id === row.id)
                                    .map((r) => Number(r.InvoiceAmount) || 0);

                                  const total = checked
                                    ? selectedRows.reduce((sum, x) => sum + x, 0)
                                    : // removing one → subtract just that row
                                    selectedRows.reduce((sum, x) => sum + x, 0) -
                                    (Number(row.InvoiceAmount) || 0);

                                  setCloseAmount(String(total));
                                }, 0);
                              }}

                            />
                          )}

                        {/* Manager sees checkmark but cannot change it */}
                        {currentUserEmail === managerEmail &&
                          row.InvoiceStatus === "Pending Close Approval" && (
                            <span className="badge bg-info">Selected</span>
                          )}
                      </td>
                    )}

                    {/* Add a checkbox for each row in the "Edit Invoice" column if the condition is satisfied */}
                    {/* {showEditInvoiceColumn &&
                      row.InvoiceStatus !== "Credit Note Uploaded"  && ( */}
                    {showEditInvoiceColumn &&
                      row.InvoiceStatus !== "Credit Note Uploaded" &&
                      row.InvoiceStatus !== "Pending Approval" &&
                      row.InvoiceStatus !== "Invoice Closed" &&
                      row.InvoiceStatus !== "Pending Manager Approval" &&
                      row.CreditNoteStatus !== "Pending" && (
                        <td className="fixedcolumn">
                          <input
                            type="checkbox"
                            className="form-check-input"
                            checked={row.invoiceApprovalChecked || false}
                            onChange={(e) =>
                              setInvoiceRows((prevRows) =>
                                prevRows.map((r) =>
                                  r.id === row.id
                                    ? {
                                        ...r,
                                        invoiceApprovalChecked:
                                          e.target.checked,
                                      }
                                    : r
                                )
                              )
                            }
                          />
                        </td>
                      )}
                    </tr>
                  );
                })}
            </tbody>
          </table>
        </div>
      </div>
      {/* </div> */}

                
{/* ================= CLOSE INVOICES SECTION ================= */}
{showCloseInvoicesSection && (
<div className="mt-4">

  {/* HEADER – matches other sections */}
  <div
    className="d-flex justify-content-between align-items-center mb-3 sectionheader"
    style={{ padding: "7px 8px" }}
  >
    <div className="d-flex align-items-center justify-content-between">
      <h5
        className="fw-bold mt-2 me-2 headingColor"
        style={{ cursor: "pointer" }}
        onClick={() => setShowCloseSection((prev) => !prev)}
        aria-expanded={showCloseSection}
        aria-controls="closeInvoicesCollapse"
      >
        Close Invoices
      </h5>
    </div>
  </div>

  {/* COLLAPSIBLE SECTION – matches others */}
  <div
    className={`${showCloseSection ? "collapse show" : "collapse"} sectioncontent`}
    id="closeInvoicesCollapse"
  >
    <div className="card card-body">



  {console.log("========= DEBUG: RAW INVOICE ROWS =========")}
{invoiceRows.forEach((r, i) => {
  console.log(
    `Row ${i}:`,
    "ID:", r.itemID,
    "| Status:", r.InvoiceStatus,
    "| CloseAmount:", r.CloseAmount,
    "| CloseReason:", r.CloseReason
  );
})}
  {console.log("============================================")}



      {/* ============ CLOSE AMOUNT + CLOSE REASON ============ */}
{isManagerButNotRequester ? (
  <>
{getPendingCloseInvoices().length > 0 && 
 getPendingCloseInvoices()[0].InvoiceStatus === "Pending Close Approval" && (
<div className="row">
  {/* Close Amount (Requested) */}
  <div className="col-md-3 mb-3">
    <label className="form-label fw-bold">Close Amount (Requested)</label>
    <div
      className="form-control"
      style={{ background: "#f3f3f3" }}
    >
      {getPendingCloseInvoices()[0]?.CloseAmount || "-"}
    </div>
  </div>

  {/* Close Reason (Requested) */}
  <div className="col-md-5 mb-3">
    <label className="form-label fw-bold">Close Reason (Requested)</label>
    <div
      className="form-control"
      style={{ background: "#f3f3f3" }}
    >
      {getPendingCloseInvoices()[0]?.CloseReason || "-"}
    </div>
  </div>
</div>
 )}
  </>
) : (
  <>
<div className="row">
  {/* Close Amount */}
  <div className="col-md-3 mb-3">
    <label className="form-label fw-bold">Close Amount</label>
    <input
      type="number"
      className="form-control"
      value={closeAmount}
      onChange={(e) => setCloseAmount(e.target.value)}
    />
  </div>

  {/* Close Reason */}
  <div className="col-md-5 mb-3">
    <label className="form-label fw-bold">Close Reason</label>
    <textarea
      className="form-control"
      rows={2}
      value={closeReason}
      onChange={(e) => setCloseReason(e.target.value)}
    />
  </div>
</div>

  </>
)}


      {/* ================== REQUESTOR VIEW ================== */}
     {currentUserEmail === requestorEmail &&
  getSelectedCloseInvoices().length > 0 &&
  !getSelectedCloseInvoices().some(
    r =>
      r.InvoiceStatus === "Pending Close Approval" ||
      r.InvoiceStatus === "Close Hold"
  ) && (
  <div className="d-flex justify-content-center gap-2 mt-3">


    {isRequestorManager ? (
      <button
        className="btn btn-danger px-4"
        style={{ width: "auto" }}
        onClick={handleDirectCloseInvoices}
      >
        Close Invoices
      </button>
    ) : (
      <button
        className="btn btn-success px-4"
        style={{ width: "auto" }}
        onClick={handleCloseApprove}
      >
        Send for Close Approval
      </button>
    )}

  </div>
)}


      {/* ================== MANAGER VIEW ================== */}

      
{currentUserEmail === managerEmail &&
 currentUserEmail !== requestorEmail &&
 invoiceRows.some(
   r =>
     r.InvoiceStatus === "Pending Close Approval"
 ) && (
  <div className="mt-3">
    <h6 className="fw-bold">Manager Actions</h6>

    <button
      className="btn btn-success me-2"
      onClick={async () => {
        const pending = getPendingCloseInvoices();
        if (pending.length === 0) {
          alert("No invoices pending close approval.");
          return;
        }
        await submitCloseApprove();

      }}
    >
      Approve Close
    </button>

   <button
  className="btn btn-danger me-2"
  onClick={() => {
    const pending = getPendingCloseInvoices();
    if (pending.length === 0) {
      alert("No pending close invoices to reject.");
      return;
    }
    setModalRow(pending[0]);   // <-- IMPORTANT
    setShowCloseRejectModal(true);
  }}
>
  Reject
</button>
    <button
      className="btn btn-warning"
      onClick={() => setShowCloseHoldModal(true)}
    >
      Hold
    </button>
  </div>
)}
    </div>
  </div>
</div>
)}


      {/* Modal for Edit */}
      <Modal
        show={showEditModal}
        onHide={handleCloseModal}
        centered
        dialogClassName="custommodalwidth" // Add custom class for width
      >
        <Modal.Header closeButton>
          <Modal.Title>Edit Invoice</Modal.Title>
        </Modal.Header>
        <Modal.Body>
          {selectedRow && (
            <FinaceInvoiceSection
              invoiceRow={selectedRow}
              siteUrl={props.siteUrl}
              context={context}
              currentUserEmail={currentUserEmail}
            />
          )}
        </Modal.Body>
        <Modal.Footer>
          <Button variant="danger" onClick={handleCloseModal}>
            Close
          </Button>
        </Modal.Footer>
      </Modal>

      {errors.invoiceTotal && (
        <div className="text-danger fw-bold mt-2">{errors.invoiceTotal}</div>
      )}

      {/* Modal for Invoice History */}
      <Modal
        show={showHistoryModal}
        onHide={() => setShowHistoryModal(false)}
        centered
        size="lg"
      >
        <Modal.Header closeButton>
          <Modal.Title>
            Invoice Payment History / Credit Note Details
          </Modal.Title>
        </Modal.Header>
        <Modal.Body>
          {historyLoading ? (
            <div className="text-center">Loading...</div>
          ) : invoiceHistoryData.length > 0 ? (
            <div className="table-responsive">
              <h5>Invoice Payment History</h5>

              <table className="table table-bordered">
                <thead className="table-light">
                  <tr>
                    <th>S.no</th>
                    <th>Invoice Tax Amount</th>
                    <th>Payment Date</th>
                    <th>Payment Amount</th>
                    <th>Pending Amount</th>
                    <th>Remarks</th>
                    <th>Financer Name</th>
                  </tr>
                </thead>
                <tbody>
                  {invoiceHistoryData.map((item, index) => (
                    <tr key={item.Id}>
                      <td>{index + 1}</td>
                      <td>{item.InvoiceTaxAmount}</td>
                      <td>
                        {item.PaymentDate
                          ? moment(item.PaymentDate).format("DD-MM-YYYY")
                          : ""}
                      </td>
                      <td>{item.PaymentAmount}</td>
                      <td>{item.PendingAmount}</td>
                      <td>{item.Comment}</td>
                      <td>{item.FinancerName || ""}</td>
                      {/* Adjust according to your SharePoint column names */}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <div className="text-center text-danger fw-bold">
              No payment received on this invoice.
            </div>
          )}

          <div className="mt-4">
            <div key={selectedRow?.itemID} className="mb-3">
              <CreditNoteDetails
                invoiceID={
                  selectedRow?.itemID ? String(selectedRow.itemID) : ""
                } // Convert to string or fallback to an empty string
                props={props}
              />
            </div>
          </div>
        </Modal.Body>
        <Modal.Footer>
          <Button variant="danger" onClick={() => setShowHistoryModal(false)}>
            Close
          </Button>
        </Modal.Footer>
      </Modal>

   {/* ================== CLOSE HOLD MODAL ================== */}
      <Modal
  show={showCloseRejectModal}
  onHide={() => setShowCloseRejectModal(false)}
  centered
>
  <Modal.Header closeButton>
    <Modal.Title>Reject Close Request</Modal.Title>
  </Modal.Header>

  <Modal.Body>
    <label className="fw-bold">Reason for Rejection</label>
    <textarea
      className="form-control"
      rows={3}
      value={managerCloseReason}
      onChange={(e) => setManagerCloseReason(e.target.value)}
    ></textarea>
  </Modal.Body>

  <Modal.Footer>
    <Button variant="secondary" onClick={() => setShowCloseRejectModal(false)}>
      Cancel
    </Button>
    <Button
  variant="danger"
  onClick={async () => {
    const selected = getSelectedCloseInvoices(); // use the existing helper
    if (selected.some((r: InvoiceRow) => r.InvoiceStatus === "Close Hold")) {
      await submitCloseHoldReject();
    } else {
      await submitCloseReject();
    }
  }}
>
  Reject Close
</Button>


  </Modal.Footer>
</Modal>

<Modal
  show={showCloseHoldModal}
  onHide={() => setShowCloseHoldModal(false)}
  centered
>
  <Modal.Header closeButton>
    <Modal.Title>Hold Close Request</Modal.Title>
  </Modal.Header>

  <Modal.Body>
    <label className="fw-bold">Reason for Hold</label>
    <textarea
      className="form-control"
      rows={3}
      value={managerCloseReason}
      onChange={(e) => setManagerCloseReason(e.target.value)}
    ></textarea>
  </Modal.Body>

  <Modal.Footer>
    <Button variant="secondary" onClick={() => setShowCloseHoldModal(false)}>
      Cancel
    </Button>
    <Button variant="warning" onClick={submitCloseHold}>
      Hold
    </Button>
  </Modal.Footer>
</Modal>


   


      {/* ================== REJECT MODAL (Bootstrap) ================== */}
      <Modal
        show={showRejectModal}
        onHide={closeModals}
        centered
      >
        <Modal.Header closeButton>
          <Modal.Title>Reject Invoice</Modal.Title>
        </Modal.Header>

        <Modal.Body>
          <label className="mt-2 fw-bold">New Invoice Due Date</label>
          <DatePicker
            format="DD-MM-YYYY"
            value={managerDueDate ? moment(managerDueDate) : null}
            onChange={(date) =>
              setManagerDueDate(date ? date.format("YYYY-MM-DD") : null)
            }
            style={{ width: "100%" }}
            disabledDate={(current) => {
              if (!modalRow?.InvoiceDueDate) return false;

              const invoiceDue = moment(modalRow.InvoiceDueDate, "DD-MM-YYYY");

              // Disable all dates BEFORE or SAME as Invoice Due Date
              return current && current <= invoiceDue.endOf("day");
            }}
            getPopupContainer={(triggerNode: HTMLElement) => triggerNode}
          />


          <label className="mt-3 fw-bold">Reason for Rejection</label>
          <textarea
            className="form-control"
            rows={3}
            value={managerReason}
            onChange={(e) => setManagerReason(e.target.value)}
          ></textarea>
        </Modal.Body>

        <Modal.Footer>
          <Button variant="secondary" onClick={closeModals}>Cancel</Button>
          <Button variant="danger" onClick={submitReject}>Reject</Button>
        </Modal.Footer>
      </Modal>

      {/* ================== HOLD MODAL (Bootstrap) ================== */}
<Modal
  show={showHoldModal}
  onHide={closeModals}
  centered
>
  <Modal.Header closeButton>
    <Modal.Title>Hold Invoice</Modal.Title>
  </Modal.Header>

  <Modal.Body>
    <label className="mt-3 fw-bold">Reason for Hold</label>
    <textarea
      className="form-control"
      rows={3}
      value={managerReason}
      onChange={(e) => setManagerReason(e.target.value)}
    ></textarea>
  </Modal.Body>

  <Modal.Footer>
    <Button variant="secondary" onClick={closeModals}>
      Cancel
    </Button>
    <Button variant="warning" onClick={submitHold}>
      Hold
    </Button>
  </Modal.Footer>
</Modal>

    </div>
  );
}
