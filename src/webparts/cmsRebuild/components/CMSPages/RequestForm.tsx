/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-expressions*/
/* eslint-disable @rushstack/no-new-null */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable prefer-const */
/* eslint-disable max-lines */
/* eslint-disable  eqeqeq */
/* eslint-disable  no-empty */
/* eslint-disable  @typescript-eslint/no-unused-vars */

import * as React from "react";

import { Snackbar, Alert } from "@mui/material";

// import { SPHttpClient } from '@microsoft/sp-http';

function getViewUrl(file: any) {
  const fileName = file.FileLeafRef || "";
  const ext = fileName.split(".").pop()?.toLowerCase() || "";
  if (ext === "txt") {
    return file.EncodedAbsUrl;
  }
  if (ext === "png" || ext === "jpg" || ext === "csv") {
    return file.EncodedAbsUrl + "?web=1";
  }
  return file.ServerRedirectedEmbedUri;
}

import { useCallback, useState, useEffect, useRef } from "react";
import "bootstrap/dist/css/bootstrap.min.css";
import RequesterInvoiceSection from "./RequesterInvoiceSection";
import { ICmsRebuildProps } from "../ICmsRebuildProps";
import moment from "moment";
// import { Form, Input, Button, Select, Spin, message, Radio, DatePicker, Upload, Space } from "antd";
import { DatePicker } from "antd";
import "./RequestForm.css";
import {
  saveDataToSharePoint,
  getSharePointData,
  updateDataToSharePoint,
  getDocumentLibraryData,
  isUserInGroup,
  addFileInSharepoint,
  handleDownload,
  deleteAttachmentFile,
} from "../services/SharePointService"; // Import the service
import { sp } from "@pnp/sp/presets/all";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
// import FinaceInvoiceSection from "./FinaceInvoiceSection";

import { Modal, Button } from "react-bootstrap"; // Import Bootstrap Modal
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import {
  faEye,
  faTrash,
  faFileArrowDown,
  faEdit,
  faPaperPlane,
  faXmark,
  faPenToSquare,
  faUpload,
  faAngleUp,
  faAngleDown,
  // faFloppyDisk,
  faBell,
} from "@fortawesome/free-solid-svg-icons";
import AzureSection from "./AzureSectiontoday";
import Dashboard from "./Dashboard"; // Import Dashboard component
// import { FaPlus, FaTrash, FaCheck } from 'react-icons/fa';
import Spinner from "react-bootstrap/Spinner";

// import Alert from "antd/es/alert/Alert";
// import { get } from "@microsoft/sp-lodash-subset";

// MilestoneBar: displays a horizontal chevron-style milestone strip
// MilestoneBar: displays a horizontal chevron-style milestone strip

// ...existing code...
const MilestoneBar: React.FC<{
  status?: string;
  creditNoteLabel?: string;
}> = ({ status, creditNoteLabel }) => {
  const stages = ["Pending From Approver", "Hold", "Reminder", "Approved"];

  // Append credit-note stage when a label is provided (e.g. "Credit Note Pending" / "Credit Note Uploaded")
  if (creditNoteLabel) {
    stages.push(creditNoteLabel);
  }

  // Determine current stage (prefer creditNoteLabel if present)
  const currentIndex = creditNoteLabel
    ? stages.indexOf(creditNoteLabel)
    : stages.indexOf(status || "");

  return (
    <div
      style={{
        display: "flex",
        alignItems: "center",
        gap: 0,
        flexWrap: "nowrap",
        overflowX: "auto",
        justifyContent: "center",
        marginTop: 20,
      }}
    >
      {stages.map((stage, i) => {
        let state: "done" | "active" | "upcoming" = "upcoming";
        if (currentIndex >= 0) {
          if (i < currentIndex) state = "done";
          else if (i === currentIndex) state = "active";
        }

        const background =
          state === "done"
            ? "linear-gradient(90deg,#28a745,#0f9d58)"
            : state === "active"
            ? "linear-gradient(90deg,#ffb84d,#ff8a00)"
            : "#e9ecef";

        const color = state === "upcoming" ? "#343a40" : "#ffffff";

        return (
          <div
            key={stage}
            title={stage}
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              padding: "10px 26px",
              color,
              fontWeight: 700,
              fontSize: 13,
              background,
              clipPath:
                "polygon(0 0, calc(100% - 18px) 0, 100% 50%, calc(100% - 18px) 100%, 0 100%, 18px 50%)",
              boxShadow:
                state === "active" ? "0 6px 18px rgba(0,0,0,0.12)" : "none",
              marginLeft: i === 0 ? 0 : -18,
              zIndex: stages.length - i,
              whiteSpace: "nowrap",
            }}
          >
            {stage}
          </div>
        );
      })}
    </div>
  );
};
// ...existing code...

// Helper: compute credit note label from selectedRow
function getCreditNoteLabel(row: any): string | undefined {
  if (!row) return undefined;
  // Only consider when IsCreditNoteUploaded flag indicates pending/upload flow
  if (row.isCreditNoteUploaded !== "No") return undefined;

  const inv = Array.isArray(row.invoiceDetails) ? row.invoiceDetails : [];

  const hasUploaded = inv.some(
    (i: any) => String(i?.CreditNoteStatus || "") === "Uploaded"
  );
  const hasPending = inv.some(
    (i: any) => String(i?.CreditNoteStatus || "") === "Pending"
  );

  if (hasUploaded) return "Credit Note Uploaded";
  if (hasPending) return "Credit Note Pending";
  // default to pending when flag is No but no details present
  return "Credit Note Pending";
}

const RequestForm = (props: ICmsRebuildProps) => {
  sp.setup({ spfxContext: { pageContext: props.context.pageContext } });

  const [snackbar, setSnackbar] = useState({
    open: false,
    message: "",
    severity: "info", // "success", "error", "warning", "info"
  });

  const showSnackbar = (
    message: string,
    severity: "success" | "error" | "warning" | "info"
  ) => {
    setSnackbar({ open: true, message, severity });
  };

  const handleCloseSnackbar = () => {
    setSnackbar({ ...snackbar, open: false });
  };

  const { context } = props;
  const [azureSectionData, setAzureSectionData] = useState<any[]>([]);
  const handleAzureSectionDataChange = (rows: any[]) => {
    setAzureSectionData(rows);
  };
  console.log(azureSectionData, "azureSectionData");

  const [siteUrl, setSiteUrl] = useState("");

  const [currentUser, setCurrentUser] = useState("");
  const [currentUserEmail, setCurrentUserEmail] = useState("");

  // useEffect(() => {
  //     if (props.) {
  //         console.log("Selected Row ID:", props.rowId); // Log the rowId in the console
  //     }
  // }, [props.rowId]);

  const [uid, setUid] = useState("");
  // const MainList = "CMSRequestNew";
  const MainList = "CMSRequest";
  const OperationalEditRequest = "OperationalCMSEditRequest";
  const OperationalEditInvoiceHistory = "OperationalCMSEditInvoiceHistory";
  // const InvoicelistName = "CMSRequestDetailsNew";
  const InvoicelistName = "CMSRequestDetails";
  const CompanyMaster = "CompanyMaster";
  const ContractTypeMaster = "ProductMaster";
  const CustomerMaster = "CustomerMaster";
  const CurrencyMaster = "CurrencyMaster";
  const AttachmentTypeMaster = "AttachmentTypeMaster";
  const ContractDocumentLibaray = "ContractDocument";
  const CMSBGDocLibaray = "CMSBGDocument";
  const CMSBGReleaseDocLibaray = "CMSBGReleaseDocument";
  const AzureSectionList = "CMSAzureList"; // Assuming this is the correct list name for Azure section
  const AzureSectionChargeList = "CMSAzureChargesList"; // Assuming this is the correct list name for Azure section
  console.log("AzureSectionList:", AzureSectionList);
  const generatedUID = Math.random().toString(36).substr(2, 16).toUpperCase(); // 16-character UID
  const [companies, setCompanies] = useState<string[]>([]);
  const [contractTypes, setContractType] = useState<string[]>([]);
  const [productServiceOptions, setProductServiceOptions] = useState<string[]>(
    []
  );
  const [allContractTypeData, setAllContractTypeData] = useState<any[]>([]);
  const [customersName, setCustomersName] = useState<string[]>([]);
  const [currencyName, setCurrencyName] = useState<string[]>([]);
  const [attachmentTypes, setAttachmentType] = useState<string[]>([]);
  const [attachmentTypesFinance, setAttachmentTypeFinance] = useState<string[]>(
    []
  );

  const [loading, setLoading] = useState(true);
  const [requestClosed, setRequestClosed] = useState<string>("");
  const [errors, setErrors] = useState<{ [key: string]: string }>({});
  const [activeUsers, setActiveUsers] = useState<string[]>([]);
  const [projectManagerUsers, setProjectManagerUsers] = useState<string[]>([]);
  const [projectLeadUsers, setProjectLeadUsers] = useState<string[]>([]);
  const [RequestId, setRequestId] = useState<string>("");

  const [file, setFile] = useState<File | null>(null);
  const [poFile, setPoFile] = useState<File | null>(null);
  const [uploadedAttachmentFiles, setAttachmentUploadedFiles] = useState<any[]>(
    []
  );
  const [poUploadedAttachmentFiles, setPoAttachmentUploadedFiles] = useState<
    any[]
  >([]);

  const [isDisabled, setIsDisabled] = useState(false);
  const [isUserInGroupM, setIsUserInGroup] = useState(false); // For CMSAccountGroup
  const [isUserInAdminM, setIsUserInAdmin] = useState(false);
  console.log(isUserInAdminM, "isUserInAdminM");
  // const [bgEndDate, setBgEndDate] = useState('');
  const [bgFile, setBGFile] = useState<File | null>(null);
  const [bgID, setBgID] = useState("");
  const [uploadedBGFiles, setUploadedBGFiles] = useState<any[]>([]);
  const [proceedButtonCount, setProceedButtonCount] = useState(0);

  console.log(proceedButtonCount);
  const [showEditBGModal, setShowEditModal] = useState(false);
  const [selectedBGReleaseFile, setSelectedBGFileDetail] = useState<any>(null);
  const [isBGRelease, setIsBGRelease] = useState("No");
  const [RelasedBGFiles, setRelasedBGFiles] = useState<any[]>([]);
  const [bgReleaseFile, setBGRelaseFile] = useState<File | null>(null);
  const [navigateToDashboard, setNavigateToDashboard] = useState(false);
  const [dashboardKey, setDashboardKey] = useState(0);
  const [isLoading, setIsLoading] = useState(false);
  const [showClientWorkDetail, setShowClientWorkDetail] = useState(true);
  const [showPODetails, setShowPODetails] = useState(false);
  const [showOtherAttachments, setShowOtherAttachments] = useState(false);
  const [showBGDetails, setShowBGDetails] = useState(false);
  const [isInvoiceSectionCollapsed, setIsInvoiceSectionCollapsed] =
    useState(true);
  const [isAzureSectionCollapsed, setIsAzureSectionCollapsed] = useState(true);
  // Approval checkboxes state for sections: client/work, PO, invoice
  const [approvalChecks, setApprovalChecks] = useState({
    client: false,
    po: false,
    invoice: false,
  });
  const poFileInputRef = useRef<HTMLInputElement>(null); // Add ref for PO file input
  const bgFileInputRef = useRef<HTMLInputElement>(null); // Add ref for PO file input
  const [isPopupOpen, setIsPopupOpen] = useState(false);
  const [reason, setReason] = useState("");
  const [isSubmitting, setIsSubmitting] = useState(false); // Loader for modal
  const [modalSnackbar, setModalSnackbar] = useState({
    open: false,
    message: "",
    severity: "info",
  });
  useEffect(() => {
    const siteUrl = props.context.pageContext.web.absoluteUrl;
    const currentUser = props.context.pageContext.user.displayName;
    const currentUserEmail = props.context.pageContext.user.email;

    setSiteUrl(siteUrl);
    setCurrentUser(currentUser);
    setCurrentUserEmail(currentUserEmail);
    // setActiveUsers([currentUserEmail]);

    console.log("Site URL:", siteUrl);
    console.log("Current User:", currentUser);
    console.log("Current User Email:", currentUserEmail);
  }, []); // Run only once on component mount

  const [deletedInvoiceItemIDs, setDeletedInvoiceItemIDs] = useState<number[]>(
    []
  );

  const [operationalEdits, setOperationalEdits] = useState<any[]>([]);
  const [loadingOperationalEdits, setLoadingOperationalEdits] = useState(false);
  const [showOperationalEdits, setShowOperationalEdits] = useState(false);
  const [showCurrentOperational, setShowCurrentOperational] = useState(true);
  const [showOldOperational, setShowOldOperational] = useState(false);
  const todayDate = new Date().toISOString().split("T")[0];
  // function formatDateForDateInput(isoDate: string | null | undefined): string {
  //   if (!isoDate) return "";
  //   const date = new Date(isoDate);
  //   const offsetDate = new Date(
  //     date.getTime() + date.getTimezoneOffset() * 60000
  //   ); // Adjust for timezone
  //   return offsetDate.toISOString().split("T")[0];
  // }

  const removeWhiteSpace = (str: string) =>
    str ? str.trim().replace(/\s+/g, " ") : "";

  function formatDateForTable(isoString: string): string {
    const date = new Date(isoString);

    const mm = (date.getMonth() + 1 < 10 ? "0" : "") + (date.getMonth() + 1);
    const dd = (date.getDate() < 10 ? "0" : "") + date.getDate();
    const yyyy = date.getFullYear();

    // return `${mm}-${dd}-${yyyy}`;
    return `${dd}-${mm}-${yyyy}`;
  }

  // Add state for customerDetails
  const [customerDetails, setCustomerDetails] = useState<any[]>([]);
  console.log(customerDetails, "customerDetails");
  useEffect(() => {
    console.log("loading", loading);

    const fetchData = async () => {
      try {
        const [
          companyData,
          contractData,
          customerData,
          currencyData,
          attachmentType,
        ] = await Promise.all([
          getSharePointData({ context }, CompanyMaster, ""),
          getSharePointData({ context }, ContractTypeMaster, ""),
          getSharePointData({ context }, CustomerMaster, ""),
          getSharePointData({ context }, CurrencyMaster, ""),
          getSharePointData({ context }, AttachmentTypeMaster, ""),
        ]);

        setCompanies(companyData.map((i) => i.Title));
        setContractType(
          Array.from(new Set(contractData.map((i) => i.ContractType)))
        );
        setAllContractTypeData(contractData); // Save full data
        console.log("contractData", contractData);
        setCustomersName(customerData.map((i) => i.Title));
        console.log(customersName, "customersName");
        setCurrencyName(currencyData.map((i) => i.Currency));
        console.log(attachmentType, "attachmentType");
        setAttachmentType(attachmentType.map((i) => i.Title));
        const financeAttachmentTypes = attachmentType
          .filter((i) => i.Finance === true) // Assuming "Finance" is the column name
          .map((i) => i.Title);
        setAttachmentTypeFinance(financeAttachmentTypes);

        // Store all 5 columns in a single state
        const customerDetailsArr = customerData.map((i: any) => ({
          Title: i.Title,
          SupportTypeRequired: i.SupportTypeRequired,
          DiscountTypeRequired: i.DiscountTypeRequired,
          Support: i.Support,
          Discount: i.Discount,
        }));
        setCustomerDetails(customerDetailsArr);
        console.log("Customer Details:", customerDetailsArr);
      } catch (error) {
        console.error("Error fetching dropdown data:", error);
      } finally {
        setLoading(false);
      }
    };

    fetchData();
  }, [context]);

  const [formData, setFormData] = useState({
    requester: "",
    requestId: "",
    requestTitle: "",
    requestType: "",
    description: "",
    requestedBy: "", // If needed
    requestDate: new Date().toISOString().split("T")[0], // Format YYYY-MM-DD
    status: "",
    department: "",
    companyName: "",
    contractType: "",
    customerName: "",
    govtContract: "",
    productServiceType: "",
    location: "",
    customerEmail: "",
    workTitle: "",
    // gstNo: "",
    workDetail: "",
    renewalRequired: "",
    renewalDate: "",
    poNo: "",
    poDate: "",
    poAmount: "",
    currency: "INR",
    approverStatus: "",
    editReason: "",
    accountManager: "",
    projectManager: "",
    attachmentType: "",
    comment: "",
    file: null,
    bgRequired: "",
    bgDate: "",
    bgEndDate: "",
    startDate: "",
    endDate: "",
    invoiceCriteria: "",
    paymentMode: "",
    poComment: "", // Added poComment
    // invoiceCri
  });

  const [invoiceRows, setInvoiceRows] = useState([
    {
      id: 1,
      InvoiceDescription: "",
      RemainingPoAmount: "",
      InvoiceAmount: "",
      InvoiceDueDate: "",
      InvoiceProceedDate: "",
      InvoiceComment: "",
      showProceed: false, // Added missing property
      InvoiceStatus: "", // Set InvoiceStatus to "Added"
      userInGroup: false, // Set InvoiceStatus to "Added"
      employeeEmail: "",
      itemID: null as number | null,
      InvoiceNo: "",
      InvoiceDate: "",
      InvoiceTaxAmount: "",
      ClaimNo: null as number | null,
      RequestID: "",
      DocId: "",
      PendingAmount: "",
      InvoiceFileID: "",
      invoiceApprovalChecked: false, // Initialize here
      invoiceCloseApprovalChecked: false, // Initialize here
      PrevInvoiceStatus: "",
      CreditNoteStatus: "",
    },
  ]);

  function getInvoiceTotal(popupRows: any[]) {
    return popupRows
      .reduce((total, row) => {
        const val = parseFloat(row.invoiceValue);
        return total + (isNaN(val) ? 0 : val);
      }, 0)
      .toFixed(2);
  }

  // const handleExit = async (event?: React.MouseEvent<HTMLButtonElement>) => {
  //   if (event && typeof event.preventDefault === "function") {
  //     event.preventDefault();
  //   }
  //   if (props.refreshCmsDetails) {
  //     await props.refreshCmsDetails();
  //   }
  //   if (props.onExit) {
  //     props.onExit();
  //     // Do not perform local navigation if onExit is provided
  //     return;
  //   }
  //   setNavigateToDashboard(true);
  //   setTimeout(() => {
  //     setDashboardKey((prev) => prev + 1);
  //     setNavigateToDashboard(true);
  //   }, 100);
  //   setNavigateToDashboard(true);
  // };

  const handleExit = async (event?: React.MouseEvent<HTMLButtonElement>) => {
    try {
      if (event && typeof event.preventDefault === "function") {
        event.preventDefault();
      }

      if (typeof props.refreshCmsDetails === "function") {
        await props.refreshCmsDetails();
      }

      if (typeof props.onExit === "function") {
        props.onExit();
        return;
      }

      console.error("onExit function is not provided.");
    } catch (error) {
      console.error("Error in handleExit:", error);
    }
  };

  // New helper: refresh UI and optionally navigate to dashboard.
  // Use navigate=false for approve/hold/reject/remind flows to avoid redirect.

  const finalizeAction = async (navigate = false) => {
    try {
      if (typeof props.refreshCmsDetails === "function") {
        await props.refreshCmsDetails();
      }

      if (typeof props.onExit === "function") {
        // call onExit but do not force local navigation if caller chose not to
        props.onExit();
        // If onExit exists we prefer it over internal navigation
        return;
      }

      if (navigate) {
        setNavigateToDashboard(true);
        setTimeout(() => {
          setDashboardKey((prev) => prev + 1);
        }, 100);
      }
    } catch (err) {
      console.error("finalizeAction error:", err);
    }
  };
  /*
  // update handleExit to reuse helper (keeps previous behaviour)
  const handleExit = async (event?: React.MouseEvent<HTMLButtonElement>) => {
    try {
      if (event && typeof event.preventDefault === "function") {
        event.preventDefault();
      }
      await finalizeAction(true); // original exit should navigate
    } catch (error) {
      console.error("Error in handleExit:", error);
    }
  };*/
  /*
  const finalizeAction = async (navigate = false) => {
  try {
    if (typeof props.refreshCmsDetails === "function") {
      await props.refreshCmsDetails();
    }

    if (typeof props.onExit === "function") {
      props.onExit();
      return; // Exit early if onExit is provided
    }

    if (navigate) {
      setNavigateToDashboard(true); // Trigger navigation
    }
  } catch (err) {
    console.error("finalizeAction error:", err);
  }
};

const handleExit = async (event?: React.MouseEvent<HTMLButtonElement>) => {
  try {
    if (event && typeof event.preventDefault === "function") {
      event.preventDefault();
    }

    // Perform cleanup before navigating
    setIsLoading(false);
    setSnackbar({ open: false, message: "", severity: "info" });

    await finalizeAction(true);
  } catch (error) {
    console.error("Error in handleExit:", error);
  }
};*/

  const getTotalInvoiceAmount = (popupRows: any[], chargeRows: any[]) => {
    const totalInvoiceValue = parseFloat(getInvoiceTotal(popupRows));
    let supportCharges = 0;
    let discountCharges = 0;

    if (!chargeRows || chargeRows.length === 0) {
      return totalInvoiceValue.toFixed(2);
    }

    chargeRows.forEach((row) => {
      const percentage = Number(row.percentage) || 0;
      const value = Number(row.value) || 0;
      let addOnValue = Number(row.calculatedValue) || 0;
      const additionalType = row.additionalType
        ? row.additionalType.toLowerCase()
        : "";
      // If you have a checkbox for additional charges, you can check it here if needed

      let total = 0;
      if (percentage) {
        if (additionalType === "percentage") {
          let totalPerc = percentage + addOnValue;
          if (totalPerc < 0) totalPerc = 0;
          total = (totalPerc * totalInvoiceValue) / 100;
        } else if (additionalType === "value") {
          total = (percentage * totalInvoiceValue) / 100 + addOnValue;
        } else {
          total = (percentage * totalInvoiceValue) / 100;
        }
      } else if (value) {
        if (additionalType === "percentage") {
          let totalPerc = addOnValue;

          if (totalPerc < 0) totalPerc = 0;
          total = (totalPerc * totalInvoiceValue) / 100 + value;
        } else if (additionalType === "value") {
          total = addOnValue + value;
        } else {
          total = value;
        }
      }

      if (row.chargesType === "Support") {
        supportCharges += total;
      } else if (row.chargesType === "Discount") {
        discountCharges += total;
      }
    });

    // Total Invoice Amount = (Total Invoice Value + Support) - Discount
    return (totalInvoiceValue + supportCharges - discountCharges).toFixed(2);
  };

  // ...existing code...

  const [editInvoiceRows, setEditInvoiceRows] = useState<any[]>([]);
  console.log(editInvoiceRows);

  // useEffect(() => {
  //   if (
  //     props.rowEdit === "Yes" &&
  //     props.selectedRow &&
  //     props.selectedRow.invoiceDetails
  //   ) {
  //     console.log(
  //       props.selectedRow.isAzureRequestClosed,
  //       "props.selectedRow.invoiceDetailsprops.selectedRow.invoiceDetails"
  //     );
  //     console.log(props.selectedRow, "props.isazuee");
  //     const invoiceData = props.selectedRow.invoiceDetails.map(
  //       (invoice: any, index: number) => ({
  //         id: index + 1,
  //         InvoiceDescription: invoice.Comments || "",
  //         RemainingPoAmount: invoice.PoAmount || "",
  //         InvoiceAmount: invoice.InvoiceAmount || "",
  //         InvoiceDueDate: invoice.InvoiceDueDate
  //           ? new Date(invoice.InvoiceDueDate).toLocaleDateString("en-GB")
  //           : "",
  //         InvoiceProceedDate: invoice.ProceedDate
  //           ? new Date(invoice.ProceedDate).toLocaleDateString("en-GB")
  //           : "",
  //         showProceed: true,
  //         InvoiceStatus: invoice.InvoiceStatus || "",
  //         userInGroup: false,
  //         employeeEmail: props.selectedRow.employeeEmail || "",
  //         InvoiceNo: invoice.InvoicNo || "",
  //         InvoiceDate: invoice.InvoiceDate || "",
  //         InvoiceTaxAmount: invoice.InvoiceTaxAmount,
  //         itemID: invoice.Id,
  //         ClaimNo: invoice.ClaimNo || null,
  //         PrevInvoiceStatus: invoice.PrevInvoiceStatus || "",
  //         RequestID: invoice.RequestID || "",
  //         DocId: invoice.DocId || "",
  //         PendingAmount: "",
  //         InvoiceFileID: invoice.InvoiceFileID,
  //       })
  //     );
  //     setEditInvoiceRows(invoiceData);
  //   }
  // }, [props.rowEdit, props.selectedRow]);

  /*
  useEffect(() => {
    if (
      props.rowEdit === "Yes" &&
      props.selectedRow &&
      props.selectedRow.invoiceDetails
    ) {
      console.log(
        props.selectedRow.isAzureRequestClosed,
        "props.selectedRow.invoiceDetailsprops.selectedRow.invoiceDetails"
      );
      console.log(props.selectedRow, "props.isazuee");

      const invoiceData = props.selectedRow.invoiceDetails.map(
        (invoice: any, index: number) => ({
          id: index + 1,
          InvoiceDescription: invoice.Comments || "",
          RemainingPoAmount: invoice.PoAmount || "",
          InvoiceAmount: invoice.InvoiceAmount || "",
          InvoiceDueDate: invoice.InvoiceDueDate
            ? new Date(invoice.InvoiceDueDate).toLocaleDateString("en-GB")
            : "",
          InvoiceProceedDate: invoice.ProceedDate
            ? new Date(invoice.ProceedDate).toLocaleDateString("en-GB")
            : "",
          showProceed: true,
          InvoiceStatus: invoice.InvoiceStatus || "",
          userInGroup: false,
          employeeEmail: props.selectedRow.employeeEmail || "",
          InvoiceNo: invoice.InvoicNo || "",
          InvoiceDate: invoice.InvoiceDate || "",
          InvoiceTaxAmount: invoice.InvoiceTaxAmount,
          itemID: invoice.Id,
          ClaimNo: invoice.ClaimNo || null,
          PrevInvoiceStatus: invoice.PrevInvoiceStatus || "",
          RequestID: invoice.RequestID || "",
          DocId: invoice.DocId || "",
          PendingAmount: "",
          InvoiceFileID: invoice.InvoiceFileID,
        })
      );

      // Update `editInvoiceRows` only if the data has changed
      setEditInvoiceRows((prevRows) => {
        const isDataChanged =
          JSON.stringify(prevRows) !== JSON.stringify(invoiceData);
        if (isDataChanged) {
          console.log("Updating editInvoiceRows with new data:", invoiceData);
          return invoiceData;
        }
        return prevRows;
      });
    }
  }, [props.rowEdit, props.selectedRow, props.selectedRow?.invoiceDetails]);
*/

  const fetchOperationalEdits = useCallback(
    async (requestId?: number) => {
      if (!requestId) {
        setOperationalEdits([]);
        return;
      }
      setLoadingOperationalEdits(true);
      try {
        const filter = `$select=*,Id&$filter=RequestID eq ${requestId}&$orderby=Id desc`;
        const data = await getSharePointData(
          { context },
          OperationalEditRequest,
          filter
        );
        setOperationalEdits(Array.isArray(data) ? data : []);
      } catch (err) {
        console.error("Failed to fetch operational edit requests:", err);
      } finally {
        setLoadingOperationalEdits(false);
      }
    },
    [context]
  );

  useEffect(() => {
    if (
      props.rowEdit === "Yes" &&
      props.selectedRow &&
      props.selectedRow.invoiceDetails
    ) {
      console.log(
        props.selectedRow.isAzureRequestClosed,
        "props.selectedRow.invoiceDetailsprops.selectedRow.invoiceDetails"
      );
      console.log(props.selectedRow, "props.isazuee");

      const invoiceData = props.selectedRow.invoiceDetails.map(
        (invoice: any, index: number) => ({
          id: index + 1,
          InvoiceDescription: invoice.Comments || "",
          RemainingPoAmount: invoice.PoAmount || "",
          InvoiceAmount: invoice.InvoiceAmount || "",
          InvoiceDueDate: invoice.InvoiceDueDate
            ? new Date(invoice.InvoiceDueDate).toLocaleDateString("en-GB")
            : "",
          InvoiceProceedDate: invoice.ProceedDate
            ? new Date(invoice.ProceedDate).toLocaleDateString("en-GB")
            : "",
          showProceed: true,
          InvoiceStatus: invoice.InvoiceStatus || "",
          userInGroup: false,
          employeeEmail: props.selectedRow.employeeEmail || "",
          InvoiceNo: invoice.InvoicNo || "",
          InvoiceDate: invoice.InvoiceDate || "",
          InvoiceTaxAmount: invoice.InvoiceTaxAmount,
          itemID: invoice.Id,
          ClaimNo: invoice.ClaimNo || null,
          PrevInvoiceStatus: invoice.PrevInvoiceStatus || "",
          CreditNoteStatus: invoice.CreditNoteStatus || "",
          RequestID: invoice.RequestID || "",
          DocId: invoice.DocId || "",
          PendingAmount: "",
          InvoiceFileID: invoice.InvoiceFileID,
        })
      );

      // Update `editInvoiceRows` only if the data has changed
      setEditInvoiceRows((prevRows) => {
        const isDataChanged =
          JSON.stringify(prevRows) !== JSON.stringify(invoiceData);
        if (isDataChanged) {
          console.log("Updating editInvoiceRows with new data:", invoiceData);
          return invoiceData;
        }
        return prevRows;
      });
      fetchOperationalEdits(props.selectedRow.id);
    }
  }, [props.rowEdit, props.selectedRow, props.selectedRow?.invoiceDetails]);

  const generateRequestId = async (
    data: { ShortName?: string; LastUsedValue?: number; Id?: number }[]
  ) => {
    if (data && data.length > 0) {
      const shortName = data[0].ShortName || "REQ";
      const nextValue = (data[0].LastUsedValue || 0) + 1;
      // const paddedValue = ('0' + nextValue).slice(-2);
      const requestID = `${shortName}/${nextValue}`;
      // setRequestId(requestID);

      const updatedata = {
        LastUsedValue: nextValue,
      };

      try {
        if (data[0].Id !== undefined) {
          const updatedData = await updateDataToSharePoint(
            CustomerMaster,
            updatedata,
            siteUrl,
            data[0].Id
          );
          console.log("updatedata", updatedData);
        } else {
          console.error("Error: data[0].Id is undefined.");
          alert("Failed to update Request ID in SharePoint due to missing ID.");
        }
      } catch (error) {
        console.error("Error updating LastUsedValue in SharePoint:", error);
        alert("Failed to update Request ID in SharePoint.");
      }

      console.log("Generated Request ID:", requestID);
      return requestID;
    }
  };

  const handleContractTypeChange = (
    e: React.ChangeEvent<HTMLSelectElement | HTMLInputElement>
  ) => {
    const { name, value } = e.target;

    setFormData((prev) => ({
      ...prev,
      [name]: value,
    }));

    // If contractType changed, update productServiceOptions
    if (name === "contractType") {
      const filtered = allContractTypeData
        .filter((item) => item.ContractType === value)
        .map((item) => item.Title);
      setInvoiceRows([
        {
          id: 1,
          InvoiceDescription: "",
          RemainingPoAmount: "",
          InvoiceAmount: "",
          InvoiceDueDate: "",
          InvoiceProceedDate: "",
          InvoiceComment: "",
          showProceed: false,
          InvoiceStatus: "",
          userInGroup: false,
          employeeEmail: "",
          itemID: null,
          InvoiceNo: "",
          InvoiceDate: "",
          InvoiceTaxAmount: "",
          ClaimNo: null as number | null,
          PrevInvoiceStatus: "",
          CreditNoteStatus: "",
          RequestID: "",
          DocId: "",
          PendingAmount: "",
          InvoiceFileID: "",
          invoiceApprovalChecked: false, // Initialize here
          invoiceCloseApprovalChecked: false, // Initialize here
        },
      ]);
      setProductServiceOptions(filtered);
      // Also reset productServiceType in form
      setFormData((prev) => ({ ...prev, productServiceType: "" }));
    }
  };
  const onPeoplePickerChange = useCallback((items: any[]): void => {
    if (items.length > 0) {
      const selectedEmails = items.map((item) => item.secondaryText);
      console.log("Selected Emails----:", selectedEmails);
      setActiveUsers(selectedEmails); // Update selected users for Account Manager
      setFormData((prev) => ({
        ...prev,
        accountManager: items[0].text || items[0].secondaryText || "",
      }));
      console.log("Updated Account Manager Picker Users:", selectedEmails);
    } else {
      console.log("Selected Emails not----:");

      setActiveUsers([]); // Clear selection
      setFormData((prev) => ({
        ...prev,
        accountManager: "",
      }));
      console.log("No user selected for Account Manager");
    }
  }, []);

  const onProjectManagerPickerChange = useCallback((items: any[]): void => {
    if (items.length > 0) {
      const selectedEmails = items.map((item) => item.secondaryText);
      setProjectManagerUsers(selectedEmails); // Update selected users for Project Manager
      setFormData((prev) => ({
        ...prev,
        projectManager: items[0].text || items[0].secondaryText || "",
      }));
      console.log("Updated Project Manager Picker Users:", selectedEmails);
    } else {
      setProjectManagerUsers([]); // Clear selection
      setFormData((prev) => ({
        ...prev,
        projectManager: "",
      }));
      console.log("No user selected for Project Manager");
    }
  }, []);

  const onProjectLeadPickerChange = useCallback((items: any[]): void => {
    if (items.length > 0) {
      const selectedEmails = items.map((item) => item.secondaryText);
      setProjectLeadUsers(selectedEmails); // Update selected users for Project Manager
      setFormData((prev) => ({
        ...prev,
        projectLead: items[0].text || items[0].secondaryText || "",
      }));
      console.log("Updated Project Manager Picker Users:", selectedEmails);
    } else {
      setProjectLeadUsers([]); // Clear selection
      setFormData((prev) => ({
        ...prev,
        projectLead: "",
      }));
      console.log("No user selected for Project Manager");
    }
  }, []);

  const validateForm = () => {
    const newErrors: { [key: string]: string } = {};

    if (!formData.requester.trim())
      newErrors.requester = "Requester is required.";
    if (!formData.companyName.trim())
      newErrors.companyName = "Company Name is required.";
    if (!formData.contractType.trim())
      newErrors.contractType = "Contract Type is required.";
    if (!formData.customerName.trim())
      newErrors.customerName = "Customer Name is required.";
    if (!formData.productServiceType.trim())
      newErrors.productServiceType = "Product/Service Type is required.";
    if (!formData.location.trim()) newErrors.location = "Location is required.";
    if (!formData.govtContract.trim())
      newErrors.govtContract = "Govt Contract selection is required."; // Add validation for Govt Contract
    // if (!formData.customerEmail) newErrors.customerEmail = "Customer Email is required.";
    if (!formData.customerEmail.trim()) {
      newErrors.customerEmail = "Customer Email is required.";
    } else {
      const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      if (!emailRegex.test(formData.customerEmail.trim())) {
        newErrors.customerEmail = "Please enter a valid customer email.";
      }
    }
    if (!activeUsers || activeUsers.length === 0) {
      newErrors.accountManager = "Account Manager is required.";
    }
    if (!projectManagerUsers || projectManagerUsers.length === 0) {
      newErrors.projectManager = "Project Manager is required.";
    }

    if (!formData.workTitle.trim())
      newErrors.workTitle = "Work Title is required.";
    // if (!formData.gstNo) newErrors.gstNo = "gst No is required.";
    if (!formData.workDetail.trim())
      newErrors.workDetail = "Work Detail is required.";
    // if (!formData.poNo.trim()) newErrors.poNo = "PO No is required.";
    // if (!formData.poDate.trim()) newErrors.poDate = "PO Date is required.";

    if (formData.bgRequired === "Yes") {
      if (!formData.poNo.trim())
        newErrors.poNo = "PO No is required when BG Required is Yes.";
      if (!formData.poDate.trim())
        newErrors.poDate = "PO Date is required when BG Required is Yes.";
    } else if (formData.poNo.trim() && !formData.poDate.trim()) {
      newErrors.poDate = "PO Date is required when PO No is filled.";
    }
    // if (!formData.poAmount.trim())
    //   newErrors.poAmount = "PO Amount is required.";
    if (
      formData.productServiceType?.toLowerCase() !== "azure" &&
      !formData.poAmount.trim()
    ) {
      newErrors.poAmount = "PO Amount is required.";
    }
    if (!formData.currency.trim()) newErrors.currency = "Currency is required.";
    if (!formData.bgRequired.trim())
      newErrors.bgRequired = "bg Required selection is required.";
    // if (!formData.bgDate) newErrors.bgDate = "BG Date is required.";

    if (formData.productServiceType === "Resource") {
      if (!formData.startDate) newErrors.startDate = "Start Date is required.";
      if (!formData.endDate) newErrors.endDate = "End Date is required.";
      if (!formData.paymentMode)
        newErrors.paymentMode = "Payment Mode is required.";
      if (!formData.invoiceCriteria.trim())
        newErrors.invoiceCriteria = "Invoice Criteria is required.";
    }

    // Validate Invoice Details

    if (formData.productServiceType?.toLowerCase() !== "azure") {
      let totalInvoiceAmount = 0;
      if (formData.productServiceType?.toLowerCase() === "resource") {
        let minRows = 1;
        switch (formData.invoiceCriteria?.toLowerCase()) {
          case "yearly":
            minRows = 1;
            break;
          case "half-yearly":
            minRows = 2;
            break;
          case "quarterly":
            minRows = 4;
            break;
          case "2-monthly":
            minRows = 6;
            break;
          default:
            minRows = 1;
        }
        if (invoiceRows.length < minRows) {
          alert(`Minimum ${minRows} invoice rows required.`);
          return false;
        }
      }
      // invoiceRows.forEach((row, index) => {
      //   if (!row.InvoiceDescription.trim())
      //     newErrors[`InvoiceDescription_${index}`] =
      //       "Invoice Description is required.";
      //   if (!row.InvoiceAmount.trim())
      //     newErrors[`InvoiceAmount_${index}`] = "Invoice Amount is required.";
      //   if (row.InvoiceAmount === "0")
      //     newErrors[`InvoiceAmount_${index}`] =
      //       "Invoice Amount can not be zero.";
      //   if (!row.InvoiceDueDate.trim())
      //     newErrors[`InvoiceDueDate_${index}`] =
      //       "Invoice Due Date is required.";
      //   totalInvoiceAmount += Number(row.InvoiceAmount) || 0;
      // });

      // if (totalInvoiceAmount !== Number(formData.poAmount.trim())) {
      //   newErrors.invoiceTotal = `Total Invoice Amount (${totalInvoiceAmount}) must equal PO Amount (${formData.poAmount}).`;
      // }

      invoiceRows.forEach((row, index) => {
        if (!row.InvoiceDescription.trim())
          newErrors[`InvoiceDescription_${index}`] =
            "Invoice Description is required.";
        if (!row.InvoiceAmount.trim())
          newErrors[`InvoiceAmount_${index}`] = "Invoice Amount is required.";
        if (row.InvoiceAmount === "0")
          newErrors[`InvoiceAmount_${index}`] =
            "Invoice Amount can not be zero.";
        if (!row.InvoiceDueDate.trim())
          newErrors[`InvoiceDueDate_${index}`] =
            "Invoice Due Date is required.";
        totalInvoiceAmount += Number(row.InvoiceAmount) || 0;
      });

      const round2 = (v: number) =>
        Math.round((Number(v) + Number.EPSILON) * 100) / 100;

      const roundedTotalInvoice = round2(totalInvoiceAmount);
      const roundedPoAmount = round2(Number(formData.poAmount.trim()) || 0);

      if (roundedTotalInvoice.toFixed(2) !== roundedPoAmount.toFixed(2)) {
        newErrors.invoiceTotal = `Total Invoice Amount (${roundedTotalInvoice.toFixed(
          2
        )}) must equal PO Amount (${roundedPoAmount.toFixed(2)}).`;
      }
    }

    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  useEffect(() => {
    const siteUrl = props.context.pageContext.web.absoluteUrl;
    const currentUser = props.context.pageContext.user.displayName;
    const currentUserEmail = props.context.pageContext.user.email;

    // setSiteUrl(siteUrl);
    // setCurrentUser(currentUser);
    // setCurrentUserEmail(currentUserEmail);
    setUid(generatedUID);

    console.log("Site URL:", siteUrl);
    console.log("Current User:", currentUser);
    console.log("Current User Email:", currentUserEmail);
    console.log("Generated UID:", generatedUID);

    setFormData((prevFormData) => ({
      ...prevFormData,
      requester: currentUser,
      requestedBy: currentUser, // If needed
    }));
  }, []);

  const handleTextFieldChange = (
    e: React.ChangeEvent<
      HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement
    >
  ) => {
    const { name, value } = e.target;

    if (
      name === "customerName" &&
      formData.productServiceType?.toLowerCase() === "azure"
    ) {
      // alert("azure");
      setAzureSectionData([
        {
          dueDate: "",
          description: "",
          isAutoGenerated: false,
          bothChargesRequired: "",
          popupRows: [],
          chargeRows: [],
        },
      ]);
      setFormData((prev) => ({
        ...prev,
        productServiceType: "",
      }));
      setInvoiceRows([
        {
          id: 1,
          InvoiceDescription: "",
          RemainingPoAmount: "",
          InvoiceAmount: "",
          InvoiceDueDate: "",
          InvoiceProceedDate: "",
          InvoiceComment: "",
          showProceed: false,
          InvoiceStatus: "",
          userInGroup: false,
          employeeEmail: "",
          itemID: null,
          InvoiceNo: "",
          InvoiceDate: "",
          InvoiceTaxAmount: "",
          ClaimNo: null,
          PrevInvoiceStatus: "",
          CreditNoteStatus: "",
          RequestID: "",
          DocId: "",
          PendingAmount: "",
          InvoiceFileID: "",
          invoiceApprovalChecked: false, // Initialize here
          invoiceCloseApprovalChecked: false, // Initialize here
        },
      ]);
      return;
    }

    if (
      name === "productServiceType" &&
      value.toLowerCase() === "azure" &&
      !formData.customerName.trim()
    ) {
      alert(
        "Please select a Customer Name before choosing Azure as Product/Service Type."
      );
      return;
    }
    if (name === "productServiceType" && value.toLowerCase() === "azure") {
      setIsAzureSectionCollapsed(true);
    }
    let updatedFormData = { ...formData, [name]: value };

    if (
      name === "productServiceType" &&
      ["amc", "resource", "azure", "", "test license"].includes(
        value.toLowerCase()
      )
    ) {
      setInvoiceRows([
        {
          id: 1,
          InvoiceDescription: "",
          RemainingPoAmount: "",
          InvoiceAmount: "",
          InvoiceDueDate: "",
          InvoiceProceedDate: "",
          InvoiceComment: "",
          showProceed: false,
          InvoiceStatus: "",
          userInGroup: false,
          employeeEmail: "",
          itemID: null,
          InvoiceNo: "",
          InvoiceDate: "",
          InvoiceTaxAmount: "",
          ClaimNo: null,
          PrevInvoiceStatus: "",
          CreditNoteStatus: "",
          RequestID: "",
          DocId: "",
          PendingAmount: "",
          InvoiceFileID: "",
          invoiceApprovalChecked: false, // Initialize here
          invoiceCloseApprovalChecked: false, // Initialize here
        },
      ]);
    }

    if (
      (name === "bgRequired" && value === "Yes" && formData.poDate) ||
      (name === "poDate" && formData.bgRequired === "Yes")
    ) {
      const poDateValue = name === "poDate" ? value : formData.poDate;
      if (poDateValue) {
        const poDateMoment = moment(poDateValue, "DD-MM-YYYY");
        if (poDateMoment.isValid()) {
          updatedFormData.bgDate = poDateMoment
            .add(20, "days")
            .format("DD-MM-YYYY");
        }
      }
    }

    setFormData(updatedFormData);
    // validateForm();
  };

  // ...existing code...
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setFile(e.target.files[0]);
    }
  };

  const uploadPOAttachment = async () => {
    setIsLoading(true);
    if (!file || !formData.attachmentType) {
      setIsLoading(false);
      // alert("Please select a Attachment Type For Po Upload.");
      // alert("Please select a Attachment Type For Po Upload.");
      alert("Please select a Attachment and Attachment Type.");
      return;
    }

    // Hide delete confirmation modal if it was open

    try {
      const updateMetadata = {
        FileID: uid,
        AttachmentType: formData.attachmentType,
        Comment: formData.comment || " ",
        RequestID: RequestId || " ",
        UserEmail: currentUserEmail || " ",
      };

      const filterQuery = `FileID eq '${uid}' and AttachmentType ne 'PO'`;
      const selectedValues =
        "*, Id, FileLeafRef, FileID, AttachmentType, Comment, FileRef, EncodedAbsUrl, ServerRedirectedEmbedUri";

      const filedata = await addFileInSharepoint(
        file,
        updateMetadata,
        ContractDocumentLibaray,
        filterQuery,
        selectedValues
      );
      console.log("context", context);
      console.log("context", filedata);

      setAttachmentUploadedFiles(filedata); // Now correctly sets the state with the returned data
      setFile(null);

      // code added by shreya on 29-May-25 fro after upload reset the fields
      setFormData((prev) => ({
        ...prev,
        attachmentType: "",
        comment: "",
      }));
      (document.getElementById("fileInputPO") as HTMLInputElement).value = "";
    } catch (error: any) {
      console.error("Error uploading file:", error);
      alert("Error uploading file.");
    }
    setIsLoading(false);
  };
  const handlePoAttachmentDownload = async (
    e: React.MouseEvent<HTMLButtonElement>,
    encodedUrl: string
  ) => {
    await handleDownload(e, encodedUrl, { context: props.context });
  };

  // view code of renuka
  const handlePoAttachmentView = (
    e: React.MouseEvent<HTMLButtonElement>,
    viewUrl: string,
    fileName: string
  ) => {
    e.preventDefault();
    let urlToOpen = viewUrl;
    if (urlToOpen) {
      window.open(urlToOpen, "_blank");
    } else {
      alert("Preview not available for this file type.");
    }
  };

  //change iew code by shreya on 29-May-25
  // const handlePoAttachmentView = (
  //   e: React.MouseEvent<HTMLButtonElement>,
  //   fileUrl: string,
  //   fileRef: string,
  //   fileName: string
  // ) => {
  //   e.preventDefault();

  //   const extension = fileName.split(".").pop()?.toLowerCase();

  //   const imageTypes = ["jpg", "jpeg", "png", "gif", "bmp", "webp"];
  //   const docViewerTypes = ["doc", "docx", "xls", "xlsx"];
  //   const openInBrowserTypes = ["pdf"]; // open directly instead of SharePoint viewer

  //   if (extension && imageTypes.includes(extension)) {
  //     window.open(fileUrl, "_blank"); // image preview
  //   } else if (extension && openInBrowserTypes.includes(extension)) {
  //     window.open(fileUrl, "_blank"); // open PDF directly
  //   } else if (extension && docViewerTypes.includes(extension)) {
  //     const viewUrl = `${props.context.pageContext.web.absoluteUrl
  //       }/_layouts/15/Doc.aspx?sourcedoc=${fileRef}&file=${encodeURIComponent(
  //         fileName
  //       )}&action=default`;
  //     window.open(viewUrl, "_blank"); // office docs
  //   } else {
  //     window.open(fileUrl, "_blank"); // fallback
  //   }
  // };

  const handlePoAttachmentDelete = async (file: any) => {
    if (window.confirm("Are you sure you want to delete this attachment?")) {
      try {
        await deleteAttachmentFile(ContractDocumentLibaray, file.Id); // Replace with actual library name

        // Hide row by removing it from the list
        const updatedFiles = uploadedAttachmentFiles.filter(
          (f) => f.Id !== file.Id
        );
        setAttachmentUploadedFiles(updatedFiles);
      } catch (error) {
        console.error("Error deleting file:", error);
        alert("Failed to delete file.");
      }
    }
  };

  const handleInvoiceChange = (
    index: number,
    field: string,
    value: string | number
  ) => {
    setInvoiceRows((prevRows) => {
      let updatedRows = [...prevRows];
      updatedRows[index] = { ...updatedRows[index], [field]: value };

      // Update RemainingPoAmount dynamically
      if (field === "InvoiceAmount") {
        const poAmt = parseFloat(formData.poAmount) || 0;

        // Filter only valid rows (exclude "Credit Note Uploaded")
        const validRows = updatedRows.filter(
          (row) => row.InvoiceStatus !== "Credit Note Uploaded"
        );

        // Start remaining amount from total PO
        let runningRemaining = poAmt;

        // Update only non–Credit Note Uploaded rows
        updatedRows = updatedRows.map((row) => {
          if (row.InvoiceStatus === "Credit Note Uploaded") {
            // Keep Credit Note Uploaded rows unchanged
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

        // Calculate totals excluding Credit Note Uploaded rows
        const totalInvoiceAmount = validRows.reduce(
          (sum, r) => sum + (Number(r.InvoiceAmount) || 0),
          0
        );
        const remainingAfter = +(poAmt - totalInvoiceAmount).toFixed(2);

        // Find last valid (non–Credit Note Uploaded) row
        const lastValidRow = [...updatedRows]
          .reverse()
          .find((row) => row.InvoiceStatus !== "Credit Note Uploaded");

        const lastRowHasValue =
          lastValidRow &&
          String(lastValidRow.InvoiceAmount).trim() !== "" &&
          Number(lastValidRow.InvoiceAmount) !== 0;

        // Add new blank row if PO not fully used and last valid row has value
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
            InvoiceComment: "",
            showProceed: false,
            InvoiceStatus: "",
            userInGroup: false,
            employeeEmail: "",
            itemID: null as number | null,
            InvoiceNo: "",
            InvoiceDate: "",
            InvoiceTaxAmount: "",
            ClaimNo: null,
            PrevInvoiceStatus: "",
            CreditNoteStatus: "",
            RequestID: "",
            DocId: "",
            PendingAmount: "",
            InvoiceFileID: "",
            invoiceApprovalChecked: false,
            invoiceCloseApprovalChecked: false,
          });
        }

        // Remove extra trailing rows when PO amount is exactly used up
        if (poAmt > 0 && remainingAfter === 0) {
          let lastFilledIndex = -1;
          for (let i = updatedRows.length - 1; i >= 0; i--) {
            const amt = String(updatedRows[i].InvoiceAmount).trim();
            if (
              amt !== "" &&
              Number(updatedRows[i].InvoiceAmount) !== 0 &&
              updatedRows[i].InvoiceStatus !== "Credit Note Uploaded"
            ) {
              lastFilledIndex = i;
              break;
            }
          }

          if (lastFilledIndex === -1) {
            // Nothing filled → keep one empty row with full PO amount
            updatedRows = [
              {
                id: updatedRows[0]?.id || 1,
                InvoiceDescription: "",
                RemainingPoAmount: poAmt.toFixed(2),
                InvoiceAmount: "",
                InvoiceDueDate: "",
                InvoiceProceedDate: "",
                InvoiceComment: "",
                showProceed: false,
                InvoiceStatus: "",
                userInGroup: false,
                employeeEmail: "",
                itemID: null as number | null,
                InvoiceNo: "",
                InvoiceDate: "",
                InvoiceTaxAmount: "",
                ClaimNo: null,
                PrevInvoiceStatus: "",
                CreditNoteStatus: "",
                RequestID: "",
                DocId: "",
                PendingAmount: "",
                InvoiceFileID: "",
                invoiceApprovalChecked: false,
                invoiceCloseApprovalChecked: false,
              },
            ];
          } else {
            // Keep only up to last valid (non-credit-note) filled row
            updatedRows = updatedRows.slice(0, lastFilledIndex + 1);
          }
        }
      }

      return updatedRows;
    });
  };

  const addInvoiceRow = () => {
    setInvoiceRows((prevRows) => {
      const totalInvoiceAmount = prevRows.reduce(
        (sum, row) => sum + (Number(row.InvoiceAmount) || 0),
        0
      );
      const remainingPoAmount =
        (Number(formData.poAmount) || 0) - totalInvoiceAmount;

      // Find max id in prevRows
      const maxId =
        prevRows.length > 0 ? Math.max(...prevRows.map((r) => r.id)) : 0;
      return [
        ...prevRows,
        {
          id: maxId + 1,
          InvoiceDescription: "",
          RemainingPoAmount: remainingPoAmount.toFixed(2),
          InvoiceAmount: "",
          InvoiceDueDate: "",
          InvoiceProceedDate: "",
          InvoiceComment: "",
          showProceed: false,
          InvoiceStatus: "",
          userInGroup: false,
          employeeEmail: "",
          itemID: null as number | null,
          InvoiceNo: "",
          InvoiceDate: "",
          InvoiceTaxAmount: "",
          ClaimNo: null as number | null,
          PrevInvoiceStatus: "",
          CreditNoteStatus: "",
          RequestID: "",
          DocId: "",
          PendingAmount: "",
          InvoiceFileID: "",
          invoiceApprovalChecked: false, // Initialize here
          invoiceCloseApprovalChecked: false, // Initialize here
        },
      ];
    });
  };

  const resetForm = () => {
    setFormData({
      requester: currentUser,
      requestId: "",
      requestTitle: "",
      requestType: "",
      description: "",
      requestedBy: currentUser,
      requestDate: new Date().toISOString().split("T")[0],
      status: "",
      department: "",
      companyName: "",
      contractType: "",
      customerName: "",
      govtContract: "",
      productServiceType: "",
      location: "",
      customerEmail: "",
      workTitle: "",
      // gstNo: "",
      workDetail: "",
      renewalRequired: "",
      renewalDate: "",
      poNo: "",
      poDate: "",
      poAmount: "",
      currency: "",
      approverStatus: "",
      editReason: "",
      accountManager: "",
      projectManager: "",
      attachmentType: "",
      comment: "",
      file: null,
      bgRequired: "",
      bgDate: "",
      bgEndDate: "",
      startDate: "",
      endDate: "",
      invoiceCriteria: "",
      paymentMode: "",
      poComment: "", // Added poComment
    });

    setInvoiceRows([
      {
        id: 1,
        InvoiceDescription: "",
        RemainingPoAmount: "",
        InvoiceAmount: "",
        InvoiceDueDate: "",
        InvoiceProceedDate: "",
        InvoiceComment: "",
        showProceed: false,
        InvoiceStatus: "",
        userInGroup: false,
        employeeEmail: "",
        itemID: null as number | null,
        InvoiceNo: "",
        InvoiceDate: "",
        InvoiceTaxAmount: "",
        ClaimNo: null,
        PrevInvoiceStatus: "",
        CreditNoteStatus: "",
        RequestID: "",
        DocId: "",
        PendingAmount: "",
        InvoiceFileID: "",
        invoiceApprovalChecked: false, // Initialize here

        invoiceCloseApprovalChecked: false, // Initialize here
      },
    ]);

    setUid(generatedUID);
    setProjectManagerUsers([]);
    setProjectLeadUsers([]);
  };
  const saveInvoiceRows = async (requestId: number) => {
    for (let index = 0; index < invoiceRows.length; index++) {
      const row = invoiceRows[index];
      const invoiceData = {
        ClaimNo: index + 1,
        PrevInvoiceStatus: row.PrevInvoiceStatus || "",
        CreditNoteStatus: row.CreditNoteStatus || "",
        DocId: uid,
        Comments: row.InvoiceDescription,
        PoAmount: Number(row.RemainingPoAmount),
        InvoiceAmount: Number(row.InvoiceAmount),
        // InvoiceDueDate: row.InvoiceDueDate
        //   ? new Date(row.InvoiceDueDate).toISOString().split("T")[0]
        //   : null,
        InvoiceDueDate: row.InvoiceDueDate
          ? moment(row.InvoiceDueDate, "DD-MM-YYYY").format("YYYY-MM-DD")
          : null,
        EmailBody: row.InvoiceComment,
        RequestID: requestId,
        InvoiceStatus: "Started",
      };

      try {
        await saveDataToSharePoint(InvoicelistName, invoiceData, siteUrl);
      } catch (error) {
        console.error(`Error saving invoice row ${index + 1}:`, error);
        alert(`Failed to save invoice row ${index + 1}.`);
      }
    }
  };

  // ...existing code...
  const saveAzureSectionData = async (requestId: number) => {
    console.log(requestId, "requestId");
    console.log(azureSectionData, "azureSectionData1");
    const filteredAzureSectionData =
      props.rowEdit === "Yes"
        ? azureSectionData.filter((row) => !row.isAutoGenerated)
        : azureSectionData;
    try {
      // Determine ClaimNo starting offset when editing an existing request
      const existingCount =
        props.rowEdit === "Yes"
          ? Number(props.selectedRow?.invoiceDetails?.length || 0)
          : 0;

      for (let i = 0; i < filteredAzureSectionData.length; i++) {
        const row = filteredAzureSectionData[i];
        const claimNo = existingCount + i + 1;

        let TotalInvoive;
        const totalInvoiceValue = getInvoiceTotal(row.popupRows || []);
        const totalInvoiceAmount = getTotalInvoiceAmount(
          row.popupRows || [],
          row.chargeRows || []
        );
        TotalInvoive = totalInvoiceValue;
        console.log(TotalInvoive);

        const cmsInvoiceData: any = {
          RequestID: requestId,
          ClaimNo: claimNo,
          PrevInvoiceStatus: row.PrevInvoiceStatus || "",
          CreditNoteStatus: row.CreditNoteStatus || "",
          InvoiceAmount: Number(totalInvoiceAmount) || 0,
          InvoiceDueDate: row.dueDate
            ? moment(row.dueDate, [
                "DD-MM-YYYY",
                "YYYY-MM-DD",
                "YYYY/MM/DD",
              ]).format("YYYY-MM-DD")
            : null,
          Comments: row.description,
          InvoiceStatus: "Started",
          DocId: uid || "",
        };

        const response1 = await saveDataToSharePoint(
          InvoicelistName,
          cmsInvoiceData,
          siteUrl
        );
        const newRequestId = response1.d.ID;

        // Submit each popup row after main data is saved
        if (row && row.popupRows && row.popupRows.length > 0) {
          for (const popupRow of row.popupRows) {
            const popupData = {
              RequestID: requestId,
              InvoiceID: newRequestId,
              AzureType: popupRow.type,
              AzureFileID: popupRow.azureFileId,
              totalInvoiceValue: Number(TotalInvoive) || 0,
              InvoiceValue: popupRow.invoiceValue,
            };

            const response2 = await saveDataToSharePoint(
              AzureSectionList,
              popupData,
              siteUrl
            );
            const newRequestAzureId = response2.d.ID;
            console.log(newRequestAzureId, AzureSectionChargeList);
          }
        }

        if (row && row.chargeRows && row.chargeRows.length > 0) {
          for (const chargeRow of row.chargeRows) {
            console.log(chargeRow, "chargeRow");

            const chargesData = {
              RequestID: requestId,
              InvoiceID: newRequestId,
              Value: Number(chargeRow.value) || 0,
              TotalCharges: Number(chargeRow.totalChargesvalue) || 0,
              Percentage: Number(chargeRow.percentage) || 0,
              ChargesType: chargeRow.chargesType,
              AddOnValue: Number(chargeRow.addOnValue) || 0,
              AdditionalChargesRequired: chargeRow.additionalChargesRequired,
              AdditionalType: chargeRow.additionalType,
            };

            const response3 = await saveDataToSharePoint(
              AzureSectionChargeList,
              chargesData,
              siteUrl
            );
            const newRequestAzureChargesId = response3.d.ID;
            console.log(newRequestAzureChargesId);
          }
        }

        // Optionally update local state to mark saved rows (set itemID/claimNo)
        setAzureSectionData((prev) => {
          const copy = Array.isArray(prev) ? [...prev] : [];
          const idx = copy.findIndex(
            (r) => r === row || r === filteredAzureSectionData[i]
          );
          if (idx !== -1) {
            copy[idx] = { ...copy[idx], itemID: newRequestId, claimNo };
          }
          return copy;
        });
      }
    } catch (error) {
      console.error("Error saving Azure section data:", error);
      throw error;
    }
  };
  // ...existing code...

  // Utility to recalculate totalChargesvalue for all azureSectionData rows before submit
  function recalcAzureSectionCharges(azureSectionRows: any[]) {
    return azureSectionRows.map((row) => {
      if (!row.chargeRows || row.chargeRows.length === 0) return row;
      const invoiceTotal = parseFloat(getInvoiceTotal(row.popupRows || []));
      const updatedChargeRows = row.chargeRows.map(
        (chargeRow: {
          calculatedValue: any;
          fromField: string;
          percentage: any;
          value: any;
        }) => {
          let addOnValue = Number(chargeRow.calculatedValue) || 0;
          let newTotal = 0;
          if (!chargeRow.fromField || chargeRow.fromField === "") {
            newTotal = 0;
          } else if (chargeRow.fromField === "Percentage") {
            const perc = Number(chargeRow.percentage) + addOnValue;
            newTotal = isNaN(perc)
              ? 0
              : parseFloat(((invoiceTotal * perc) / 100).toFixed(2));
          } else if (chargeRow.fromField === "Value") {
            const val = Number(chargeRow.value) + addOnValue;
            newTotal = isNaN(val) ? 0 : val;
          }
          return { ...chargeRow, totalChargesvalue: newTotal };
        }
      );
      return { ...row, chargeRows: updatedChargeRows };
    });
  }

  const handleSubmit = async (event: any) => {
    event.preventDefault();
    if (props.rowEdit !== "Yes") {
      setIsLoading(true);
      // Force recalc of all totalChargesvalue fields before validation and submit
      if (formData.productServiceType?.toLowerCase() === "azure") {
        setAzureSectionData((prev) => recalcAzureSectionCharges(prev));
      }
      // Find the earliest due date in azureSectionData
      const earliestDueDate = azureSectionData
        .map((row) => row.dueDate)
        .filter((date) => !!date)
        .map((date) => moment(date, ["DD-MM-YYYY", "YYYY-MM-DD"]))
        .filter((m) => m.isValid())
        .sort((a, b) => a.valueOf() - b.valueOf())[0];

      let nextMayFirst = null;
      if (earliestDueDate) {
        // If earliestDueDate is before April 1 of that year, use that year; else, use next year
        const year =
          earliestDueDate.month() < 3 // months are 0-indexed, so 3 is April
            ? earliestDueDate.year()
            : earliestDueDate.year() + 1;
        nextMayFirst = moment(`01-05-${year}`, "DD-MM-YYYY");
        // Example usage:
        console.log("Next 1st April:", nextMayFirst.format("DD-MM-YYYY"));
      }
      const nextMayFirstStr = nextMayFirst
        ? nextMayFirst.format("YYYY-MM-DD")
        : null;
      try {
        if (!validateForm()) {
          alert("Please fix the errors in the form.");
          setShowPODetails(true);
          setShowOtherAttachments(true);
          setShowBGDetails(true);
          setIsInvoiceSectionCollapsed(true);

          return;
        }
        if (formData.poNo.trim() || formData.bgRequired === "Yes") {
          if (poUploadedAttachmentFiles.length === 0) {
            alert("Please upload PO attachment.");
            setShowPODetails(true);
            setShowOtherAttachments(true);
            setShowBGDetails(true);
            setIsInvoiceSectionCollapsed(true);

            return;
          }
        }

        if (formData.accountManager === "") {
          alert("Please select Account Manager.");
          setShowPODetails(true);
          setShowOtherAttachments(true);
          setShowBGDetails(true);
          setIsInvoiceSectionCollapsed(true);

          return;
        }
        if (formData.projectManager === "") {
          alert("Please select Project Manager.");
          setShowPODetails(true);
          setShowOtherAttachments(true);
          setShowBGDetails(true);
          setIsInvoiceSectionCollapsed(true);

          return;
        }
        if (activeUsers.length === 0) {
          alert("Please select at least one Account Manager.");
          setShowPODetails(true);
          setShowOtherAttachments(true);
          setShowBGDetails(true);
          setIsInvoiceSectionCollapsed(true);

          return;
        }
        if (projectManagerUsers.length === 0) {
          alert("Please select at least one Project Manager.");
          setShowPODetails(true);
          setShowOtherAttachments(true);
          setShowBGDetails(true);
          setIsInvoiceSectionCollapsed(true);

          return;
        }

        const poNoFilterQuery = `$select=PoNo,CompanyName&$filter=CustomerName eq '${encodeURIComponent(
          formData.customerName
        )}' and PoNo eq '${encodeURIComponent(formData.poNo)}'`;
        const poNoData = await getSharePointData(
          { context },
          MainList,
          poNoFilterQuery
        );

        if (poNoData && poNoData.length > 0) {
          alert(
            "A record with this PO No already exists for the selected Customer Name."
          );
          setShowPODetails(true);
          setShowOtherAttachments(true);
          setShowBGDetails(true);
          setIsInvoiceSectionCollapsed(true);

          return;
        }
        // Resolve user IDs
        const userPromises = activeUsers.map(async (email) => {
          const user = await sp.web.ensureUser(`i:0#.f|membership|${email}`);
          return user.data.Id;
        });
        const projectManagerUserPromises = projectManagerUsers.map(
          async (email) => {
            const user = await sp.web.ensureUser(`i:0#.f|membership|${email}`);
            return user.data.Id;
          }
        );
        const projectLeadPromises = projectLeadUsers.map(async (email) => {
          const user = await sp.web.ensureUser(`i:0#.f|membership|${email}`);
          return user.data.Id;
        });
        const userIds = await Promise.all(userPromises);
        const projectLeadUserIds = await Promise.all(projectLeadPromises);
        const projectManagerUserIds = await Promise.all(
          projectManagerUserPromises
        );
        if (!userIds.length) {
          alert("Please select at least one Account Manager.");
          setShowPODetails(true);
          setShowOtherAttachments(true);
          setShowBGDetails(true);
          setIsInvoiceSectionCollapsed(true);

          return;
        }
        if (!projectManagerUserIds.length) {
          alert("Please select at least one Project Manager.");
          setShowPODetails(true);
          setShowOtherAttachments(true);
          setShowBGDetails(true);
          setIsInvoiceSectionCollapsed(true);

          return;
        }

        if (formData.productServiceType?.toLowerCase() === "azure") {
          // if (azureSectionData.length < 6) {
          //   setIsLoading(false);
          //   alert("There are minimum 6 rows in Azure section.");

          //   return;
          // }
          for (const row of azureSectionData) {
            const totalInvoiceValue = getInvoiceTotal(row.popupRows || []);

            const totalInvoiceAmount = getTotalInvoiceAmount(
              row.popupRows || [],
              row.chargeRows || []
            );

            if (!row.description.trim()) {
              alert("Please fill Description in Azure section.");
              setShowPODetails(true);
              setShowOtherAttachments(true);
              setShowBGDetails(true);
              setIsInvoiceSectionCollapsed(true);
              return;
            }

            if (!row.popupRows || row.popupRows.length === 0) {
              alert("Please fill additional row details in Azure section.");
              setShowPODetails(true);
              setShowOtherAttachments(true);
              setShowBGDetails(true);
              setIsInvoiceSectionCollapsed(true);

              return;
            }

            // Validate popupRows
            for (const popupRow of row.popupRows) {
              if (!popupRow.type.trim() || !popupRow.invoiceValue.trim()) {
                alert(
                  "Type and Invoice Value are required in Azure Section for each row."
                );
                setShowPODetails(true);
                setShowOtherAttachments(true);
                setShowBGDetails(true);
                setIsInvoiceSectionCollapsed(true);

                return;
              }

              if (
                !popupRow.azureFileId && // or use azureFileName/azureFileUrl as needed
                !popupRow.azureFileName &&
                !popupRow.azureFileUrl
              ) {
                setIsLoading(false);
                alert(
                  "File upload is mandatory in Azure section for each row."
                );
                setShowPODetails(true);
                setShowOtherAttachments(true);
                setShowBGDetails(true);
                setIsInvoiceSectionCollapsed(true);

                return;
              }
            }

            if (!row.chargeRows || row.chargeRows.length === 0) {
              alert("Please fill Charges Details in Azure Section.");
              setShowPODetails(true);
              setShowOtherAttachments(true);
              setShowBGDetails(true);
              setIsInvoiceSectionCollapsed(true);

              return;
            }

            if (row.chargeRows && row.chargeRows.length > 0) {
              for (const chargeRow of row.chargeRows) {
                // if (chargeRow.totalChargesvalue === 0) {
                //   alert("Please fill charges details properly.");
                //   setShowPODetails(true);
                //   setShowOtherAttachments(true);
                //   setShowBGDetails(true);
                //   setIsInvoiceSectionCollapsed(true);

                //   return;
                // }
                if (Number(chargeRow.totalChargesvalue) < 0) {
                  setIsLoading(false);
                  alert(
                    "Total Charges value cannot be negative in Azure section."
                  );
                  setShowPODetails(true);
                  setShowOtherAttachments(true);
                  setShowBGDetails(true);
                  setIsInvoiceSectionCollapsed(true);

                  return;
                }

                if (chargeRow.additionalChargesRequired === "Yes") {
                  if (
                    !chargeRow.additionalType ||
                    chargeRow.additionalType.trim() === ""
                  ) {
                    setIsLoading(false);
                    alert(
                      "Additional Type is mandatory when Additional Charges Required is Yes in Azure section."
                    );
                    setShowPODetails(true);
                    setShowOtherAttachments(true);
                    setShowBGDetails(true);
                    setIsInvoiceSectionCollapsed(true);

                    return;
                  }

                  if (
                    chargeRow.addOnValue === undefined ||
                    chargeRow.addOnValue === null ||
                    chargeRow.addOnValue === "" ||
                    chargeRow.addOnValue === 0
                  ) {
                    setIsLoading(false);
                    alert(
                      "Add On Value is mandatory when Additional Charges Required is Yes in Azure section."
                    );

                    setShowPODetails(true);
                    setShowOtherAttachments(true);
                    setShowBGDetails(true);
                    setIsInvoiceSectionCollapsed(true);

                    return;
                  }
                }
              }
            }
            if (Number(totalInvoiceValue) <= 0) {
              alert("Total Invoice Value must be greater than 0.");
              setShowPODetails(true);
              setShowOtherAttachments(true);
              setShowBGDetails(true);
              setIsInvoiceSectionCollapsed(true);

              return;
            }
            if (Number(totalInvoiceAmount) <= 0) {
              alert("Total Invoice Amount must be greater than 0.");
              setShowPODetails(true);
              setShowOtherAttachments(true);
              setShowBGDetails(true);
              setIsInvoiceSectionCollapsed(true);

              return;
            }
          }
        }
        const filterQuery = `$select=*&$filter=Title eq '${encodeURIComponent(
          formData.customerName
        )}'&$orderby=Id desc`;
        const data = await getSharePointData(
          { context },
          CustomerMaster,
          filterQuery
        );
        console.log("Request ID data", data);

        const requestID = await generateRequestId(data);
        console.log("Generated Request ID:", requestID);

        const requestData = {
          EmployeeName: formData.requester,
          UID: uid,
          EmployeeEmail: currentUserEmail,
          AzureCloseDate: nextMayFirstStr,
          RequestID: requestID,
          CompanyName: formData.companyName,
          ContractType: formData.contractType,
          CustomerName: formData.customerName,
          ProductType: formData.productServiceType,
          Location: formData.location,
          GovtContract: formData.govtContract,
          BGRequired: formData.bgRequired,
          CustomerEmail: formData.customerEmail,
          WorkTitle: formData.workTitle,

          WorkDetails: formData.workDetail,
          PoNo: formData.poNo,

          PoDate:
            formData.poDate && formData.poDate.trim()
              ? moment(formData.poDate, "DD-MM-YYYY").format("YYYY-MM-DD")
              : null,
          BGDate:
            formData.bgDate && formData.bgDate.trim()
              ? moment(formData.bgDate, "DD-MM-YYYY").format("YYYY-MM-DD")
              : null,
          StartDateResource:
            formData.startDate && formData.startDate.trim()
              ? moment(formData.startDate, "DD-MM-YYYY").format("YYYY-MM-DD")
              : null,
          EndDateResource:
            formData.endDate && formData.endDate.trim()
              ? moment(formData.endDate, "DD-MM-YYYY").format("YYYY-MM-DD")
              : null,
          InvoiceCriteria: formData.invoiceCriteria || null,
          PaymentMode: formData.paymentMode || null,

          POAmount: Number(formData.poAmount),
          Currency: formData.currency,
          RunWF: "Yes",
          AccountMangerId: userIds[0],
          ProjectManagerId: projectManagerUserIds[0],
          ProjectLeadId: projectLeadUserIds[0],
        };

        const mainListResponse = await saveDataToSharePoint(
          MainList,
          requestData,
          siteUrl
        );
        console.log("Main List Response:", mainListResponse);
        console.log("Main List Response:", formData.productServiceType);

        if (formData.productServiceType?.toLowerCase() === "azure") {
          if (mainListResponse && mainListResponse.d.ID) {
            await saveAzureSectionData(mainListResponse.d.ID);
          }
        } else {
          if (mainListResponse && mainListResponse.d.ID) {
            await saveInvoiceRows(mainListResponse.d.ID);
          }
        }

        resetForm();
        await props.refreshCmsDetails();
        setIsLoading(false);
        alert("Form and data submitted successfully!");
        if (props.onExit) {
          props.onExit();
          return;
        }
        setNavigateToDashboard(true);
        setTimeout(() => {
          setDashboardKey((prev) => prev + 1);
          setNavigateToDashboard(true);
        }, 100);
      } catch (error: any) {
        console.error("Error submitting form:", error);

        if (error.message) {
          console.error("Error Message:", error.message);
        }
        if (error.stack) {
          console.error("Error Stack:", error.stack);
        }
        if (error.response) {
          console.error("Error Response:", error.response);
        }

        alert(
          `Error submitting form: ${
            error.message || "An unknown error occurred."
          }`
        );
      } finally {
        setIsLoading(false);
      }
    } else {
      if (formData.productServiceType?.toLowerCase() === "azure") {
        await saveAzureSectionData(props.selectedRow?.id);
      }
    }
  };

  // const handleUpdatePO = async (
  //   e: React.MouseEvent<HTMLButtonElement, MouseEvent>,
  //   id: number
  // ) => {
  //   e.preventDefault();
  //   setIsLoading(true);
  //   console.log("update button clicked", id);
  //   if (!formData.poNo.trim() || !formData.poDate) {
  //     setIsLoading(false);
  //     alert("Please fill in the PO Number and PO Date before updating.");
  //     setShowPODetails(true);

  //     setIsInvoiceSectionCollapsed(true);

  //     return;
  //   }

  //   if (formData.poNo.trim() || formData.bgRequired === "Yes") {
  //     if (poUploadedAttachmentFiles.length === 0) {
  //       setIsLoading(false);
  //       alert("Please upload PO attachment.");
  //       setShowPODetails(true);
  //       setShowOtherAttachments(true);
  //       setShowBGDetails(true);
  //       setIsInvoiceSectionCollapsed(true);

  //       return;
  //     }
  //   }

  //   const poNoFilterQuery = `$select=PoNo,CompanyName&$filter=CustomerName eq '${encodeURIComponent(
  //     formData.customerName
  //   )}' and PoNo eq '${encodeURIComponent(formData.poNo)}'`;
  //   const poNoData = await getSharePointData(
  //     { context },
  //     MainList,
  //     poNoFilterQuery
  //   );

  //   if (poNoData && poNoData.length > 0) {
  //     setShowPODetails(true);
  //     setShowOtherAttachments(true);
  //     setShowBGDetails(true);
  //     setIsInvoiceSectionCollapsed(true);

  //     setIsLoading(false);
  //     alert(
  //       "A record with this PO No already exists for the selected Customer Name."
  //     );
  //     return;
  //   }

  //   const updatedata = {
  //     PoNo: formData.poNo,
  //     PoDate:
  //       formData.poDate && formData.poDate.trim()
  //         ? moment(formData.poDate, "DD-MM-YYYY").format("YYYY-MM-DD")
  //         : null,
  //     RunWF: "Yes",
  //   };

  //   try {
  //     const updatedData = await updateDataToSharePoint(
  //       MainList,
  //       updatedata,
  //       siteUrl,
  //       id
  //     );
  //     console.log("updatedata", updatedData);

  //     await props.refreshCmsDetails();
  //     setIsLoading(false);
  //     alert("Po Request has been updated successfully.");
  //     if (props.onExit) {
  //       props.onExit();
  //       return;
  //     }
  //     setNavigateToDashboard(true);
  //     setTimeout(() => {
  //       setDashboardKey((prev) => prev + 1);
  //       setNavigateToDashboard(true);
  //     }, 100);
  //     setNavigateToDashboard(true);
  //   } catch (error) {
  //     console.error("Failed to update request:", error);
  //     alert("Something went wrong while Update the data. Please try again.");
  //   } finally {
  //     setIsLoading(false);
  //   }
  // };
  // const handleProceedReject = async (
  //   e: React.MouseEvent<HTMLButtonElement, MouseEvent>,
  //   id: number
  // ) => {
  //   e.preventDefault();
  //   setIsLoading(true);
  //   console.log("update button clicked", id);
  //   if (!formData.poNo.trim() || !formData.poDate) {
  //     setIsLoading(false);
  //     alert("Please fill in the PO Number and PO Date before updating.");
  //     setShowPODetails(true);

  //     setIsInvoiceSectionCollapsed(true);

  //     return;
  //   }

  //   if (formData.poNo.trim() || formData.bgRequired === "Yes") {
  //     if (poUploadedAttachmentFiles.length === 0) {
  //       setIsLoading(false);
  //       alert("Please upload PO attachment.");
  //       setShowPODetails(true);
  //       setShowOtherAttachments(true);
  //       setShowBGDetails(true);
  //       setIsInvoiceSectionCollapsed(true);

  //       return;
  //     }
  //   }

  //   const poNoFilterQuery = `$select=PoNo,CompanyName&$filter=CustomerName eq '${encodeURIComponent(
  //     formData.customerName
  //   )}' and PoNo eq '${encodeURIComponent(formData.poNo)}'`;
  //   const poNoData = await getSharePointData(
  //     { context },
  //     MainList,
  //     poNoFilterQuery
  //   );

  //   if (poNoData && poNoData.length > 0) {
  //     setShowPODetails(true);
  //     setShowOtherAttachments(true);
  //     setShowBGDetails(true);
  //     setIsInvoiceSectionCollapsed(true);

  //     setIsLoading(false);
  //     alert(
  //       "A record with this PO No already exists for the selected Customer Name."
  //     );
  //     return;
  //   }

  //   const updatedata = {
  //     PoNo: formData.poNo,
  //     PoDate:
  //       formData.poDate && formData.poDate.trim()
  //         ? moment(formData.poDate, "DD-MM-YYYY").format("YYYY-MM-DD")
  //         : null,
  //     RunWF: "Yes",
  //   };

  //   try {
  //     const updatedData = await updateDataToSharePoint(
  //       MainList,
  //       updatedata,
  //       siteUrl,
  //       id
  //     );
  //     console.log("updatedata", updatedData);

  //     await props.refreshCmsDetails();
  //     setIsLoading(false);
  //     alert("Po Request has been updated successfully.");
  //     if (props.onExit) {
  //       props.onExit();
  //       return;
  //     }
  //     setNavigateToDashboard(true);
  //     setTimeout(() => {
  //       setDashboardKey((prev) => prev + 1);
  //       setNavigateToDashboard(true);
  //     }, 100);
  //     setNavigateToDashboard(true);
  //   } catch (error) {
  //     console.error("Failed to update request:", error);
  //     alert("Something went wrong while Update the data. Please try again.");
  //   } finally {
  //     setIsLoading(false);
  //   }
  // };
  // const handleProceedApprove = async (
  //   e: React.MouseEvent<HTMLButtonElement, MouseEvent>,
  //   id: number
  // ) => {
  //   e.preventDefault();
  //   setIsLoading(true);
  //   console.log("update button clicked", id);
  //   if (!formData.poNo.trim() || !formData.poDate) {
  //     setIsLoading(false);
  //     alert("Please fill in the PO Number and PO Date before updating.");
  //     setShowPODetails(true);

  //     setIsInvoiceSectionCollapsed(true);

  //     return;
  //   }

  //   if (formData.poNo.trim() || formData.bgRequired === "Yes") {
  //     if (poUploadedAttachmentFiles.length === 0) {
  //       setIsLoading(false);
  //       alert("Please upload PO attachment.");
  //       setShowPODetails(true);
  //       setShowOtherAttachments(true);
  //       setShowBGDetails(true);
  //       setIsInvoiceSectionCollapsed(true);

  //       return;
  //     }
  //   }

  //   const poNoFilterQuery = `$select=PoNo,CompanyName&$filter=CustomerName eq '${encodeURIComponent(
  //     formData.customerName
  //   )}' and PoNo eq '${encodeURIComponent(formData.poNo)}'`;
  //   const poNoData = await getSharePointData(
  //     { context },
  //     MainList,
  //     poNoFilterQuery
  //   );

  //   if (poNoData && poNoData.length > 0) {
  //     setShowPODetails(true);
  //     setShowOtherAttachments(true);
  //     setShowBGDetails(true);
  //     setIsInvoiceSectionCollapsed(true);

  //     setIsLoading(false);
  //     alert(
  //       "A record with this PO No already exists for the selected Customer Name."
  //     );
  //     return;
  //   }

  //   const updatedata = {
  //     PoNo: formData.poNo,
  //     PoDate:
  //       formData.poDate && formData.poDate.trim()
  //         ? moment(formData.poDate, "DD-MM-YYYY").format("YYYY-MM-DD")
  //         : null,
  //     RunWF: "Yes",
  //   };

  //   try {
  //     const updatedData = await updateDataToSharePoint(
  //       MainList,
  //       updatedata,
  //       siteUrl,
  //       id
  //     );
  //     console.log("updatedata", updatedData);

  //     await props.refreshCmsDetails();
  //     setIsLoading(false);
  //     alert("Po Request has been updated successfully.");
  //     if (props.onExit) {
  //       props.onExit();
  //       return;
  //     }
  //     setNavigateToDashboard(true);
  //     setTimeout(() => {
  //       setDashboardKey((prev) => prev + 1);
  //       setNavigateToDashboard(true);
  //     }, 100);
  //     setNavigateToDashboard(true);
  //   } catch (error) {
  //     console.error("Failed to update request:", error);
  //     alert("Something went wrong while Update the data. Please try again.");
  //   } finally {
  //     setIsLoading(false);
  //   }
  // };

  // Loader overlay
  const LoaderOverlay = () => (
    <div
      style={{
        position: "fixed",
        top: 0,
        left: 0,
        width: "100vw",
        height: "100vh",
        background: "rgba(255,255,255,0.6)",
        zIndex: 100000,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
      }}
    >
      <Spinner animation="border" variant="primary" />
      <span className="ms-3">Processing...</span>
    </div>
  );

  useEffect(() => {
    const siteUrl = props.context.pageContext.web.absoluteUrl;

    const fetchData = async () => {
      if (props.rowEdit === "Yes" && props.selectedRow) {
        console.log(
          props.selectedRow,
          "",
          currentUserEmail,
          props.context.pageContext.user.email
        );
        console.log(props.selectedRow.isPaymentReceived, "8888");
        console.log(props.selectedRow, "88818");
        setRequestClosed(props.selectedRow?.isPaymentReceived || "No");
        console.log(
          "Populating form with selected row data:",
          props.selectedRow
        );
        setIsDisabled(true); // Disable form fields in edit mode
        const selectedContractType = props.selectedRow.contractType || "";
        const selectedProductType = props.selectedRow.productType || "";
        // const startDt = props.selectedRow.startDate ? moment(props.selectedRow.startDate, "DD/MM/YYYY").format("DD-MM-YYYY") : "";
        const startDt = props.selectedRow.startDate
          ? moment(props.selectedRow.startDate, moment.ISO_8601, true).format(
              "DD-MM-YYYY"
            )
          : "";
        const endDt = props.selectedRow.endDate
          ? moment(props.selectedRow.endDate, moment.ISO_8601, true).format(
              "DD-MM-YYYY"
            )
          : "";
        console.log(
          "Start Date:",
          startDt,
          "Start Date",
          props.selectedRow.startDate,
          "End Date:",
          endDt,
          "End Date:",
          props.selectedRow.endDate
        );
        console.log("Selected Contract ID:", props.selectedRow.contractNo);
        // Always set both picker arrays and form fields together
        const accountManagerEmail = props.selectedRow.accountMangerEmail || "";
        const projectManagerEmail = props.selectedRow.projectMangerEmail || "";
        const projectLeadEmail = props.selectedRow.projectLeadEmail || "";
        console.log("Account Manager Email:", accountManagerEmail);
        console.log("Project Manager Email:", projectManagerEmail);

        setActiveUsers(accountManagerEmail ? [accountManagerEmail] : []);
        setProjectManagerUsers(
          projectManagerEmail ? [projectManagerEmail] : []
        );

        setProjectLeadUsers(projectManagerEmail ? [projectManagerEmail] : []);
        console.log(props.selectedRow.paymentMode, "paymentMode");
        setFormData((prevFormData) => ({
          ...prevFormData,
          requestId: props.selectedRow.contractNo || "",
          customerName: props.selectedRow.customerName || "",
          productServiceType: "",
          poNo: props.selectedRow.poNo || "",
          workTitle: props.selectedRow.workTitle || "",
          upcomingInvoice: props.selectedRow.upcomingInvoice || "",
          taxInvoiceAmount: props.selectedRow.taxInvoiceAmount || 0,
          editReason: props.selectedRow.editReason || "",
          totalPaymentRecievedAmt:
            props.selectedRow.totalPaymentRecievedAmt || 0,
          totalPendingAmt: props.selectedRow.totalPendingAmt || 0,
          requester: props.selectedRow.employeeName || "",
          status: "",
          accountMangerId: props.selectedRow.accountMangerId || "",

          companyName: props.selectedRow.companyName || "",
          contractType: selectedContractType, // Ensure contractType is set
          govtContract: props.selectedRow.govtContract || "No", // Ensure govtContract is set
          bgRequired: props.selectedRow.bgRequired || "No", // Ensure bgRequired is set
          location: props.selectedRow.location || "",
          customerEmail: props.selectedRow.customerEmail || "",
          workDetail: props.selectedRow.workDetail || "",
          currency: props.selectedRow.currency || "",
          approverStatus: props.selectedRow.approverStatus || "",
          // poDate: formatDateForDateInput(props.selectedRow.poDate) || "",
          poDate: props.selectedRow.poDate
            ? moment(props.selectedRow.poDate, "DD/MM/YYYY").format(
                "DD-MM-YYYY"
              )
            : "",
          poAmount: props.selectedRow.poAmount || "",
          attachmentType: "",
          comment: "",
          file: null,
          // bgDate: formatDateForDateInput(props.selectedRow.bgDate) || "",
          // bgDate:
          //   new Date(props.selectedRow.bgDate).toLocaleDateString("en-GB") ||
          //   "",
          bgDate: props.selectedRow.bgDate
            ? moment(props.selectedRow.bgDate).format("DD-MM-YYYY")
            : "",
          bgEndDate: "",
          startDate: startDt,
          endDate: endDt,
          invoiceCriteria: props.selectedRow.invoiceCriteria || "",
          paymentMode: props.selectedRow.paymentMode || "",
          // --- Ensure these are set for edit mode ---
          accountManager: accountManagerEmail,
          projectManager: projectManagerEmail,
          projectLead: projectLeadEmail,
        }));

        const UserInGroup = await isUserInGroup("CMSAccountGroup");
        const UserInAdmin = await isUserInGroup("CMSAdminGroup");
        console.log(UserInGroup, "UserInGroup");
        console.log("UserInAdmin", UserInAdmin);
        setIsUserInGroup(UserInGroup); // Save CMSAccountGroup status
        setIsUserInAdmin(UserInAdmin);
        // 2. Wait for allContractTypeData to populate (can skip if already populated)
        const filteredOptions = allContractTypeData
          .filter((item) => item.ContractType === selectedContractType)
          .map((item) => item.Title);

        // 3. Set dropdown options
        setProductServiceOptions(filteredOptions);

        // 4. Finally, set the actual productServiceType value if it exists in options
        if (filteredOptions.includes(selectedProductType)) {
          setFormData((prevFormData) => ({
            ...prevFormData,
            productServiceType: selectedProductType,
          }));
        } else {
          console.warn(
            "Product type not found in options:",
            selectedProductType
          );
        }

        // Show alert if user is in the groupIs BG Release
        if (UserInGroup) {
          // alert("User is in CMSAccountGroup");
        } else if (UserInAdmin) {
          // alert("User is in CMSAdminGroup");
        }

        setUid(props.selectedRow.docID); // Set UID from selected row
        setRequestId(props.selectedRow.contractNo);
        console.log(props.selectedRow, "props.selectedRow.docID");

        const filterQuery = `$select=*,Id,FileLeafRef,FileID,AttachmentType,Comment,FileRef,EncodedAbsUrl,ServerRedirectedEmbedUri&$filter=(FileID eq '${props.selectedRow.docID}' and AttachmentType ne 'PO')&$orderby=Id desc`;

        const updatedData = await getDocumentLibraryData(
          ContractDocumentLibaray,
          filterQuery,
          siteUrl
        );
        console.log("updatedata", updatedData);
        setAttachmentUploadedFiles(updatedData);

        const bgFilterQuery = `$select=*,Id,FileLeafRef,FileID,BGDate,BGID,IsBGRelease,FileRef,EncodedAbsUrl,ServerRedirectedEmbedUri&$filter=FileID eq '${props.selectedRow.docID}'&$orderby=Id desc`;
        const fetchedBgData = await getDocumentLibraryData(
          CMSBGDocLibaray,
          bgFilterQuery,
          siteUrl
        );
        console.log("fetchedBgData", fetchedBgData);
        setUploadedBGFiles(fetchedBgData);

        const poFilterQuery = `$select=*,Id,FileLeafRef,FileID,AttachmentType,Comment,FileRef,EncodedAbsUrl,ServerRedirectedEmbedUri&$filter=(FileID eq '${props.selectedRow.docID}' and AttachmentType eq 'PO')&$orderby=Id desc`;

        const updatedPoData = await getDocumentLibraryData(
          ContractDocumentLibaray,
          poFilterQuery,
          siteUrl
        );
        console.log("updatedata", updatedPoData);
        setPoAttachmentUploadedFiles(updatedPoData);

        // Dynamically append invoice rows
        if (props.selectedRow.invoiceDetails) {
          let invoiceData = props.selectedRow.invoiceDetails.map(
            (invoice: any, index: number) => ({
              id: index + 1,
              InvoiceDescription: invoice.Comments || "",
              RemainingPoAmount: invoice.PoAmount || "",
              InvoiceAmount: invoice.InvoiceAmount || "",

              // InvoiceDueDate:formatDateForDateInput(invoice.InvoiceDueDate) || "",

              InvoiceDueDate:
                new Date(invoice.InvoiceDueDate).toLocaleDateString("en-GB") ||
                "",
              // InvoiceProceedDate:
              //   invoice.ProceedDate || new Date().toISOString().split("T")[0],
              InvoiceProceedDate:
                new Date(invoice.ProceedDate).toLocaleDateString("en-GB") || "",
              // poDate: props.selectedRow.poDate
              //   ? moment(props.selectedRow.poDate, "DD/MM/YYYY").format(
              //     "DD-MM-YYYY"
              //   )
              //   : "",
              showProceed: true, // Show Proceed button
              InvoiceStatus: invoice.InvoiceStatus || "",

              userInGroup: UserInGroup, // Add userInGroup property
              employeeEmail: props.selectedRow.employeeEmail || "",
              InvoiceNo: invoice.InvoicNo || "",
              InvoiceDate: invoice.InvoiceDate || "",
              InvoiceTaxAmount: invoice.InvoiceTaxAmount,
              itemID: invoice.Id,
              ClaimNo: invoice.ClaimNo || "",
              PrevInvoiceStatus: invoice.PrevInvoiceStatus || "",
              CreditNoteStatus: invoice.CreditNoteStatus || "",
              RequestID: invoice.RequestID || "",
              DocId: invoice.DocId || "",
              PendingAmount: "",
              InvoiceFileID: "",
              //emailBody: invoice.EmailBody || "",
              // addOnAmount: invoice.AddonAmountValue || "",
            })
          );
          //Add conditiom for CMSAccountGroup users for only proceeded invoices by Shreya on 29-May-25
          // ✅ Apply filter for CMSAccountGroup users
          if (UserInGroup && !UserInAdmin) {
            invoiceData = invoiceData.filter(
              (invoice: any) => invoice.InvoiceStatus !== "Started"
            );
          }
          setInvoiceRows(invoiceData);
          console.log("invoiceData", invoiceData);
        }

        if (
          props.rowEdit === "Yes" &&
          props.selectedRow &&
          props.selectedRow.invoiceDetails
        ) {
          const invoiceData = props.selectedRow.invoiceDetails.map(
            (invoice: any, index: number) => ({
              id: index + 1,
              InvoiceDescription: invoice.Comments || "",
              RemainingPoAmount: invoice.PoAmount || "",
              InvoiceAmount: invoice.InvoiceAmount || "",
              InvoiceDueDate: invoice.InvoiceDueDate
                ? new Date(invoice.InvoiceDueDate).toLocaleDateString("en-GB")
                : "",
              InvoiceProceedDate: invoice.ProceedDate
                ? new Date(invoice.ProceedDate).toLocaleDateString("en-GB")
                : "",
              showProceed: true,
              InvoiceStatus: invoice.InvoiceStatus || "",
              userInGroup: false,
              employeeEmail: props.selectedRow.employeeEmail || "",
              InvoiceNo: invoice.InvoicNo || "",
              InvoiceDate: invoice.InvoiceDate || "",
              InvoiceTaxAmount: invoice.InvoiceTaxAmount,
              itemID: invoice.Id,
              ClaimNo: invoice.ClaimNo || null,
              PrevInvoiceStatus: invoice.PrevInvoiceStatus || "",
              CreditNoteStatus: invoice.CreditNoteStatus || "",
              RequestID: invoice.RequestID || "",
              DocId: invoice.DocId || "",
              PendingAmount: "",
              InvoiceFileID: invoice.InvoiceFileID,
            })
          );
          setInvoiceRows(invoiceData);
        }
        //   if (
        //     props.rowEdit === "Yes" &&
        //     (props.selectedRow.productType).toLowerCase() === "azure"
        //   ) {
        //     if (props.selectedRow.invoiceDetails) {
        //       console.log(props.selectedRow.invoiceDetails, "invoiceDetails");

        //       const azureData = await fetchAzureData(props.selectedRow.id);
        //       console.log("Azure Data:", azureData);
        //     }
        //   }
      }
    };

    fetchData(); // call the async function
  }, [props.rowEdit, props.selectedRow, allContractTypeData]);

  useEffect(() => {
    console.log(context);
    setBGFile(null);
    setBgID(generatedUID);
  }, []);

  const UploadBgFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setBGFile(e.target.files[0]);
    }
  };

  const UploadBgFile = async (e: { preventDefault: () => void }) => {
    e.preventDefault();
    setIsLoading(true);
    // Validation checks
    if (!bgFile) {
      setIsLoading(false);
      alert("Please select a BG file to upload.");
      return;
    }

    if (!formData.bgEndDate) {
      setIsLoading(false);
      alert("Please select BG End Date.");
      return;
    }

    //BGDate: formData.bgDate
    // ? moment(formData.bgDate, "DD-MM-YYYY").format("YYYY-MM-DD")
    // : null,
    try {
      // Prepare metadata
      const updateMetadata = {
        FileID: uid,
        // BGDate: formData.bgEndDate,
        BGDate: formData.bgEndDate
          ? moment(formData.bgEndDate, "DD-MM-YYYY").format("YYYY-MM-DD")
          : null,
        RequestID: RequestId || " ",
        IsBGRelease: "No",
        BGID: bgID,
        UserEmail: currentUserEmail || " ",
      };

      const filterQuery = `FileID eq '${uid}'`;
      const selectedValues =
        "Id,FileLeafRef,FileID,AttachmentType,FileRef,EncodedAbsUrl,ServerRedirectedEmbedUri,IsBGRelease,BGDate,BGID,UserEmail";

      // Upload file
      const fileData = await addFileInSharepoint(
        bgFile,
        updateMetadata,
        CMSBGDocLibaray,
        filterQuery,
        selectedValues
      );

      console.log("File upload result:", fileData);

      // Update state
      setUploadedBGFiles(fileData);
      setBGFile(null);
      setBgID(generatedUID);
      setFormData((prev) => ({
        ...prev,
        bgEndDate: "", // or use null if you prefer
      }));
      if (bgFileInputRef.current) {
        bgFileInputRef.current.value = "";
      }
    } catch (error: any) {
      console.error("Error uploading file:", error);
      setIsLoading(false);
      alert("Error uploading file. Please try again.");
    }
    setIsLoading(false);
  };
  const handleViewBGFile = (
    e: React.MouseEvent<HTMLButtonElement>,
    viewUrl: string,
    fileName: string
  ) => {
    e.preventDefault();

    let urlToOpen = viewUrl;
    if (urlToOpen) {
      window.open(urlToOpen, "_blank");
    } else {
      alert("Preview not available for this file type.");
    }
  };

  const handlePoView = (
    e: React.MouseEvent<HTMLButtonElement>,
    viewUrl: string,
    fileName: string
  ) => {
    e.preventDefault();

    let urlToOpen = viewUrl;
    if (urlToOpen) {
      window.open(urlToOpen, "_blank");
    } else {
      alert("Preview not available for this file type.");
    }
  };

  const handleBgDownload = async (
    e: React.MouseEvent<HTMLButtonElement>,
    encodedUrl: string
  ) => {
    await handleDownload(e, encodedUrl, { context: props.context });
  };

  const handleDeleteReleasedBGFile = async (
    e: React.MouseEvent<HTMLButtonElement, MouseEvent>,
    file: { Id: number }
  ) => {
    e.preventDefault();
    if (
      window.confirm("Are you sure you want to delete this released BG file?")
    ) {
      try {
        await deleteAttachmentFile(CMSBGReleaseDocLibaray, file.Id); // Replace with actual library name

        // Remove from UI
        const updatedFiles = RelasedBGFiles.filter((f) => f.Id !== file.Id);
        setRelasedBGFiles(updatedFiles);
      } catch (error) {
        console.error("Error deleting released BG file:", error);
        alert("Failed to delete the released BG file.");
      }
    }
  };

  const handleEditClick = async (e: any, bgFile: any) => {
    e.preventDefault();
    setSelectedBGFileDetail(bgFile);
    // Always fetch BG release data for the selected item
    try {
      const bgReleaseFilterQuery = `$select=Id,FileLeafRef,FileID,DocID,BGReleaseDate,FileRef,EncodedAbsUrl,ServerRedirectedEmbedUri,UserEmail&$filter=FileID eq '${uid}' and DocID eq '${bgFile.BGID}'&$orderby=Id desc`;
      const fetchedBgData = await getDocumentLibraryData(
        CMSBGReleaseDocLibaray,
        bgReleaseFilterQuery,
        siteUrl
      );
      if (fetchedBgData && fetchedBgData.length > 0) {
        setIsBGRelease("Yes");
        setRelasedBGFiles(fetchedBgData);
      } else {
        setIsBGRelease("No");
        setRelasedBGFiles([]);
      }
    } catch (error) {
      setIsBGRelease("No");
      setRelasedBGFiles([]);
      console.error("Error fetching BG Release documents:", error);
    }
    setShowEditModal(true);
  };

  const handleCloseBGModal = () => {
    setShowEditModal(false);
    setSelectedBGFileDetail(null);
    setRelasedBGFiles([]);
  };

  const handleDeleteBGFile = async (
    e: React.MouseEvent<HTMLButtonElement, MouseEvent>,
    file: { Id: number }
  ) => {
    e.preventDefault();
    if (window.confirm("Are you sure you want to delete this BG file?")) {
      try {
        await deleteAttachmentFile(CMSBGDocLibaray, file.Id); // ← Use actual library name here

        // Remove file from UI
        const updatedFiles = uploadedBGFiles.filter((f) => f.Id !== file.Id);
        setUploadedBGFiles(updatedFiles);
      } catch (error) {
        console.error("Error deleting BG file:", error);
        alert("Failed to delete BG file.");
      }
    }
  };

  const handleBGReleaseChange = async (
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    const value = e.target.value;
    setIsBGRelease(value); // Updates the radio state

    if (value === "Yes") {
      try {
        const bgReleaseFilterQuery = `$select=Id,FileLeafRef,FileID,DocID,BGReleaseDate,FileRef,EncodedAbsUrl,ServerRedirectedEmbedUri,UserEmail&$filter=FileID eq '${uid}' and DocID eq '${selectedBGReleaseFile.BGID}'&$orderby=Id desc`;

        const fetchedBgData = await getDocumentLibraryData(
          CMSBGReleaseDocLibaray,
          bgReleaseFilterQuery,
          siteUrl
        );

        console.log("fetchedBgData", fetchedBgData);
        setRelasedBGFiles(fetchedBgData);
      } catch (error) {
        console.error("Error fetching BG Release documents:", error);
        // alert("Error fetching BG Release documents.");
      }
    }
  };

  const UploadBgRelaseFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setBGRelaseFile(e.target.files[0]);
    }
  };
  const handleBGRelaseFileUpload = async (e: {
    preventDefault: () => void;
  }) => {
    e.preventDefault();
    setIsLoading(true);

    // Validation checks
    if (!bgReleaseFile) {
      setIsLoading(false);
      alert("Please select a BG file to upload.");

      return;
    }

    try {
      // Prepare metadata
      const updateMetadata = {
        FileID: uid,
        DocID: selectedBGReleaseFile.BGID,
        BGReleaseDate: new Date().toISOString().split("T")[0],
        RequestID: RequestId || " ",
        UserEmail: currentUserEmail || " ",
      };

      const filterQuery = `FileID eq '${uid}' and DocID eq '${selectedBGReleaseFile.BGID}'`;
      const selectedValues =
        "Id,FileLeafRef,FileID,DocID,BGReleaseDate,FileRef,EncodedAbsUrl,ServerRedirectedEmbedUri,UserEmail";

      // Upload file
      const fileData = await addFileInSharepoint(
        bgReleaseFile,
        updateMetadata,
        CMSBGReleaseDocLibaray,
        filterQuery,
        selectedValues
      );

      console.log("File upload result:", fileData);

      // Update state
      setRelasedBGFiles(fileData);
      // setSelectedBGFileDetail(null);
      //setBgEndDate(""); // Clear BG end date
      setBGRelaseFile(null); // Clear file input
      setIsLoading(false);
    } catch (error: any) {
      console.error("Error uploading file:", error);

      alert("Error uploading file. Please try again.");
      setIsLoading(false);
    }
  };

  // const getVersionHistory = async () => {
  //   try {
  //     // Get the item ID from props.selectedRow (adjust if needed)
  //     const itemId = props.selectedRow?.idprops.selectedRow?.id;
  //     if (!itemId) {
  //       alert("No item selected for version history.");
  //       return;
  //     }

  //     const siteUrl = props.context.pageContext.web.absoluteUrl;
  //     const listName = "CMSRequest";
  //     // REST API endpoint for item versions
  //     const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemId})/versions`;

  //     const response = await props.context.spHttpClient.get(
  //       endpoint,
  //       SPHttpClient.configurations.v1,
  //       {
  //         headers: {
  //           'Accept': 'application/json;odata=nometadata',
  //           'odata-version': ''
  //         }
  //       }
  //     );

  //     if (response.ok) {
  //       const data = await response.json();
  //       console.log("Version History:", data.value);
  //       alert(`Found ${data.value.length} versions. See console for details.`);
  //       // You can display the versions in a modal or table as needed
  //     } else {
  //       alert("Failed to fetch version history.");
  //     }
  //   } catch (error) {
  //     console.error("Error fetching version history:", error);
  //     alert("Error fetching version history.");
  //   }
  // };

  // const handleInvoiceCriteriaChange = (
  //   e: React.ChangeEvent<HTMLSelectElement>
  // ) => {
  //   const value = e.target.value;
  //   handleTextFieldChange(e);

  //   if (formData.productServiceType?.toLowerCase() === "resource" && value) {
  //     let criteriaRows = 1;
  //     switch (value.toLowerCase()) {
  //       case "yearly":
  //         criteriaRows = 1;
  //         break;
  //       case "half-yearly":
  //         criteriaRows = 2;
  //         break;
  //       case "quarterly":
  //         criteriaRows = 4;
  //         break;
  //       case "2-monthly":
  //         criteriaRows = 6;
  //         break;
  //       default:
  //         criteriaRows = 1;
  //     }
  //     if (
  //       criteriaRows > 0 &&
  //       formData.startDate &&
  //       formData.endDate
  //     ) {
  //       const totalPoAmount = Number(formData.poAmount) || 0;
  //       const dividedAmount = Math.floor((totalPoAmount / criteriaRows) * 100) / 100;
  //       let remaining = totalPoAmount;

  //       const start = moment(formData.startDate, "DD-MM-YYYY");
  //       const end = moment(formData.endDate, "DD-MM-YYYY");
  //       const totalMonths = end.diff(start, "months", true);
  //       const periodMonths = totalMonths / criteriaRows;
  //     const totalNoOfRows = Math.ceil(totalMonths / criteriaRows);

  //       console.log(totalMonths,criteriaRows,totalNoOfRows)

  //       const newRows = Array.from({ length: totalNoOfRows }, (_, index) => {
  //         let amount;
  //         if (index === criteriaRows - 1) {
  //           amount = remaining.toFixed(2);
  //         } else {
  //           amount = dividedAmount.toFixed(2);
  //           remaining -= dividedAmount;
  //         }

  //         let dueDate;
  //         if (formData.paymentMode?.toLowerCase() === "pre") {
  //           // Pre-payment: due date is the start of each period
  //           dueDate = start.clone().add(periodMonths * index, "months");
  //         } else {
  //           // Post-payment: due date is the last day of the period + 1 day
  //           let periodEnd = start.clone().add(periodMonths * (index + 1), "months").subtract(1, "days");
  //           dueDate = periodEnd.clone().add(1, "days");
  //         }

  //         return {
  //           id: index + 1,
  //           InvoiceDescription: "",
  //           RemainingPoAmount: (totalPoAmount - dividedAmount * index).toFixed(2),
  //           InvoiceAmount: amount,
  //           InvoiceDueDate: dueDate.format("DD-MM-YYYY"),
  //           InvoiceProceedDate: "",
  //           InvoiceComment: "",
  //           showProceed: false,
  //           InvoiceStatus: "",
  //           userInGroup: false,
  //           employeeEmail: "",
  //           itemID: null as number | null,
  //           InvoiceNo: "",
  //           InvoiceDate: "",
  //           InvoiceTaxAmount: "",
  //           ClaimNo: "",
  //           RequestID: "",
  //           DocId: "",
  //           PendingAmount: "",
  //           InvoiceFileID: "",
  //         };
  //       });

  //       setInvoiceRows(newRows);
  //     }
  //   }
  // };

  const handleInvoiceCriteriaChange = (
    e: React.ChangeEvent<HTMLSelectElement>
  ) => {
    const value = e.target.value;
    handleTextFieldChange(e);

    if (formData.productServiceType?.toLowerCase() === "resource" && value) {
      if (formData.startDate && formData.endDate) {
        const start = moment(formData.startDate, "DD-MM-YYYY");
        const end = moment(formData.endDate, "DD-MM-YYYY");
        const totalMonths = end.diff(start, "months", true); // fractional months

        let criteriaMonths = 1;
        switch (value.toLowerCase()) {
          case "yearly":
            criteriaMonths = 12;
            break;
          case "half-yearly":
            criteriaMonths = 6;
            break;
          case "quarterly":
            criteriaMonths = 3;
            break;
          case "2-monthly":
            criteriaMonths = 2;
            break;
          case "monthly":
            criteriaMonths = 1;
            break;
          default:
            criteriaMonths = 12;
        }

        // Calculate number of rows based on period
        const totalNoOfRows = Math.ceil(totalMonths / criteriaMonths);

        if (totalNoOfRows > 0) {
          const totalPoAmount = Number(formData.poAmount) || 0;
          const dividedAmount =
            Math.floor((totalPoAmount / totalNoOfRows) * 100) / 100;
          let remaining = totalPoAmount;

          const newRows = Array.from({ length: totalNoOfRows }, (_, index) => {
            let amount;
            if (index === totalNoOfRows - 1) {
              amount = remaining.toFixed(2);
            } else {
              amount = dividedAmount.toFixed(2);
              remaining -= dividedAmount;
            }

            // Calculate period start and end
            const periodStart = start
              .clone()
              .add(criteriaMonths * index, "months");
            let periodEnd = periodStart
              .clone()
              .add(criteriaMonths, "months")
              .subtract(1, "days");
            // Ensure last period ends at the actual end date
            if (index === totalNoOfRows - 1) {
              periodEnd = end.clone();
            }

            let dueDate;
            if (formData.paymentMode?.toLowerCase() === "pre") {
              // Pre-payment: first day of period
              dueDate = periodStart;
            } else {
              // Post-payment: last day of period + 1 day
              dueDate = periodEnd.clone().add(1, "days");
            }

            return {
              id: index + 1,
              InvoiceDescription: "",
              RemainingPoAmount: (
                totalPoAmount -
                dividedAmount * index
              ).toFixed(2),
              InvoiceAmount: amount,
              InvoiceDueDate: dueDate.format("DD-MM-YYYY"),
              InvoiceProceedDate: "",
              InvoiceComment: "",
              showProceed: false,
              InvoiceStatus: "",
              userInGroup: false,
              employeeEmail: "",
              itemID: null as number | null,
              InvoiceNo: "",
              InvoiceDate: "",
              InvoiceTaxAmount: "",
              ClaimNo: null,
              PrevInvoiceStatus: "",
              CreditNoteStatus: "",
              RequestID: "",
              DocId: "",
              PendingAmount: "",
              InvoiceFileID: "",
              invoiceApprovalChecked: false, // Initialize here
              invoiceCloseApprovalChecked: false, // Initialize here
            };
          });

          setInvoiceRows(newRows);
        }
      }
    }
  };

  const uploadPO = async () => {
    setIsLoading(true);
    if (!poFile) {
      setIsLoading(false);
      // alert("Please select file to upload.");

      alert("Please select a Attachment ");
      return;
    }

    if (!formData.poNo.trim() || !formData.poDate.trim()) {
      setIsLoading(false);
      alert(
        "Please fill in the PO Number and PO Date before uploading the PO file."
      );
      return;
    }

    try {
      const updateMetadata = {
        FileID: uid,
        AttachmentType: "PO",
        Comment: formData.poComment || " ",
        RequestID: RequestId || " ",
        UserEmail: currentUserEmail || " ",
      };

      const filterQuery = `FileID eq '${uid}' and AttachmentType eq 'PO'`;

      const selectedValues =
        "*, Id, FileLeafRef, FileID, AttachmentType,Comment, FileRef, EncodedAbsUrl,ServerRedirectedEmbedUri";

      const filedata = await addFileInSharepoint(
        poFile,
        updateMetadata,
        ContractDocumentLibaray,
        filterQuery,
        selectedValues
      );
      console.log("context", context);
      console.log("context", filedata);

      setPoAttachmentUploadedFiles(filedata);
      setPoFile(null);
      setFormData((prev) => ({ ...prev, poComment: "" }));
      // Use ref to clear file input
      if (poFileInputRef.current) {
        poFileInputRef.current.value = "";
      }
    } catch (error: any) {
      console.error("Error uploading file:", error);
      // alert("Error uploading file.");
    }
    setIsLoading(false);
  };

  const handlePoFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setPoFile(e.target.files[0]);
    }
  };

  const handlePoDelete = async (file: any) => {
    if (window.confirm("Are you sure you want to delete this attachment?")) {
      try {
        await deleteAttachmentFile(ContractDocumentLibaray, file.Id); // Replace with actual library name

        // Hide row by removing it from the list
        const updatedFiles = poUploadedAttachmentFiles.filter(
          (f) => f.Id !== file.Id
        );
        setPoAttachmentUploadedFiles(updatedFiles);
      } catch (error) {
        console.error("Error deleting file:", error);
        alert("Failed to delete file.");
      }
    }
  };

  const handlePoDownload = async (
    e: React.MouseEvent<HTMLButtonElement>,
    encodedUrl: string
  ) => {
    await handleDownload(e, encodedUrl, { context: props.context });
  };

  // if (navigateToDashboard) {
  //   return (
  //     <Dashboard
  //       key={dashboardKey}
  //       description={props.description}
  //       context={props.context}
  //       siteUrl={props.siteUrl}
  //       userGroups={props.userGroups}
  //       cmsDetails={props.cmsDetails}
  //       refreshCmsDetails={props.refreshCmsDetails}
  //       selectedMenu="Home"
  //     />
  //   );
  // }

  if (navigateToDashboard) {
    if (!props || !props.context || !props.cmsDetails) {
      console.error("Missing required props for Dashboard component.");
      return null; // Prevent rendering if props are invalid
    }

    return (
      <Dashboard
        key={dashboardKey}
        description={props.description}
        context={props.context}
        siteUrl={props.siteUrl}
        userGroups={props.userGroups}
        cmsDetails={props.cmsDetails}
        refreshCmsDetails={props.refreshCmsDetails}
        selectedMenu="Home"
      />
    );
  }

  // Only show Edit Request button if at least one row has showProceed true
  // const hasProceedButton = invoiceRows.some(row => row.showProceed);

  // Submit handler to save only the last Azure row when it does not have an ItemID
  const handleSubmitAzureDetails = async (
    e: React.MouseEvent<HTMLButtonElement>
  ) => {
    try {
      e.preventDefault?.();
    } catch {}

    if (!azureSectionData || azureSectionData.length === 0) {
      alert("No Azure rows to submit.");
      return;
    }

    const lastIndex = azureSectionData.length - 1;
    const lastRow = azureSectionData[lastIndex];
    if (!lastRow) {
      alert("No Azure row found.");
      return;
    }

    if (lastRow.itemID) {
      alert("Last Azure row already saved. Nothing to submit.");
      return;
    }

    if (!lastRow.description || !String(lastRow.description).trim()) {
      alert("Please fill Description in Azure section.");
      return;
    }
    if (!lastRow.popupRows || lastRow.popupRows.length === 0) {
      alert("Please add at least one file/type row in Azure section.");
      return;
    }
    if (!lastRow.chargeRows || lastRow.chargeRows.length === 0) {
      alert("Please add Charges Details in Azure section.");
      return;
    }

    setIsLoading(true);
    try {
      const parentRequestId =
        props.rowEdit === "Yes"
          ? props.selectedRow?.id
          : RequestId
          ? Number(RequestId)
          : null;

      if (!parentRequestId) {
        alert(
          "Parent request id is not available. Please save main request first."
        );
        return;
      }

      // compute totals
      const totalInvoiceValue = getInvoiceTotal(lastRow.popupRows || []);
      const totalInvoiceAmount = getTotalInvoiceAmount(
        lastRow.popupRows || [],
        lastRow.chargeRows || []
      );

      // compute ClaimNo (SNo)
      let claimNo = lastIndex + 1;
      if (props.rowEdit === "Yes" && props.selectedRow?.invoiceDetails) {
        // If editing existing request use existing invoice count + 1 to avoid collisions
        claimNo = (props.selectedRow.invoiceDetails?.length || 0) + 1;
      }

      const cmsInvoiceData: any = {
        RequestID: parentRequestId,
        ClaimNo: claimNo,

        InvoiceAmount: Number(totalInvoiceAmount) || 0,
        InvoiceDueDate: lastRow.dueDate
          ? moment(lastRow.dueDate, [
              "DD-MM-YYYY",
              "YYYY-MM-DD",
              "YYYY/MM/DD",
            ]).format("YYYY-MM-DD")
          : null,
        Comments: lastRow.description || "",
        InvoiceStatus: "Started",
        DocId: uid || "", // keep linkage if needed
      };

      const response1 = await saveDataToSharePoint(
        InvoicelistName,
        cmsInvoiceData,
        siteUrl
      );
      const savedInvoiceId = response1?.d?.ID;
      if (!savedInvoiceId) {
        throw new Error("Failed to save invoice main row.");
      }

      // save popup rows (AzureSectionList)
      if (lastRow.popupRows && lastRow.popupRows.length > 0) {
        for (const popupRow of lastRow.popupRows) {
          const popupData: any = {
            RequestID: parentRequestId,
            InvoiceID: savedInvoiceId,
            AzureType: popupRow.type || "",
            AzureFileID: popupRow.azureFileId || popupRow.azureFileID || "",
            totalInvoiceValue: Number(totalInvoiceValue) || 0,
            InvoiceValue: Number(popupRow.invoiceValue) || 0,
          };
          // eslint-disable-next-line no-await-in-loop
          await saveDataToSharePoint(AzureSectionList, popupData, siteUrl);
        }
      }

      // save charge rows (AzureSectionChargeList)
      if (lastRow.chargeRows && lastRow.chargeRows.length > 0) {
        for (const chargeRow of lastRow.chargeRows) {
          const chargesData: any = {
            RequestID: parentRequestId,
            InvoiceID: savedInvoiceId,
            Value: Number(chargeRow.value) || 0,
            TotalCharges: Number(chargeRow.totalChargesvalue) || 0,
            Percentage: Number(chargeRow.percentage) || 0,
            ChargesType: chargeRow.chargesType || "",
            AddOnValue: Number(chargeRow.addOnValue) || 0,
            AdditionalChargesRequired:
              chargeRow.additionalChargesRequired || "No",
            AdditionalType: chargeRow.additionalType || "",
          };
          // eslint-disable-next-line no-await-in-loop
          await saveDataToSharePoint(
            AzureSectionChargeList,
            chargesData,
            siteUrl
          );
        }
      }

      // update local state to mark saved row and store claimNo
      setAzureSectionData((prev) => {
        const copy = Array.isArray(prev) ? [...prev] : [];
        if (copy[lastIndex]) {
          copy[lastIndex] = {
            ...copy[lastIndex],
            itemID: savedInvoiceId,
            claimNo: claimNo,
          };
        }
        return copy;
      });

      resetForm();
      await props.refreshCmsDetails?.();
      setIsLoading(false);

      alert("Azure details submitted successfully.");
      if (props.onExit) {
        props.onExit();
        return;
      }
      setNavigateToDashboard(true);
      setTimeout(() => {
        setDashboardKey((prev) => prev + 1);
        setNavigateToDashboard(true);
      }, 100);
    } catch (error) {
      console.error("Error saving Azure section row:", error);
      alert("Failed to save Azure details. See console for error.");
    } finally {
      setIsLoading(false);
    }
  };

  //------------- Edit Functionality
  const [selectedId, setSelectedId] = useState<number | null>(null);

  const handleEditRequestApproval = async (
    e: React.MouseEvent<HTMLButtonElement, MouseEvent>,
    id: number
  ) => {
    e.preventDefault();
    console.log("update button clicked", id);

    if (
      !approvalChecks.client &&
      !approvalChecks.po &&
      !approvalChecks.invoice
    ) {
      return;
    }
    if (
      approvalChecks.invoice && // Use approvalChecks.invoice instead
      !invoiceRows.some((row) => row.invoiceApprovalChecked)
    ) {
      showSnackbar("Please select at least one invoice.", "error");
      return;
    }

    if (approvalChecks.invoice) {
      const selectedInvoices = invoiceRows.filter(
        (row) => row.invoiceApprovalChecked
      );
      console.log("Selected Invoice Details:", selectedInvoices);
    }

    // Log other necessary information
    console.log("Approval Checks:", approvalChecks);
    console.log("Selected ID:", id);

    setSelectedId(id); // Set the selected ID for the popup
    setIsPopupOpen(true); // Open the popup

    // if (!formData.editReason || !formData.editReason.trim()) {
    //   alert("Edit Reason is required.");
    //   return;
    // }
  };

  const handleClosePopup = () => {
    setIsPopupOpen(false); // Close the popup
    setReason(""); // Reset the reason field
  };
  const handleSubmitEditRequestApproval = async (
    event?: React.MouseEvent<HTMLButtonElement, MouseEvent>,
    id?: number
  ) => {
    try {
      event?.preventDefault();
      if (typeof id !== "number") {
        setModalSnackbar({
          open: true,
          message: "No request selected for approval.",
          severity: "error",
        });
        return;
      }

      if (!reason || reason.trim() === "") {
        setModalSnackbar({
          open: true,
          message: "Please provide a reason for the edit request.",
          severity: "error",
        });
        return;
      }

      setIsSubmitting(true); // Show loader

      const selectedSections = Object.entries(approvalChecks)
        .filter(([_, checked]) => checked)
        .map(([section]) => section)
        .join(", ")
        .trim();

      let selectedInvoiceIDs: any[] = [];
      if (approvalChecks.invoice) {
        selectedInvoiceIDs = invoiceRows
          .filter((row) => row.invoiceApprovalChecked)
          .map((row) => row.itemID)
          .filter((id) => id !== null);
      }

      const editRequestData = {
        RequestID: id,
        Reason: reason.trim(),
        UserName: currentUser,
        UserEmail: currentUserEmail,
        SelectedSections: selectedSections,
        Status: "Pending Approval",
        InvoiceID: selectedInvoiceIDs.join(", "),
        ContractID: formData.requestId || "",
        ReminderDate: todayDate,
      };

      await saveDataToSharePoint(
        OperationalEditRequest,
        editRequestData,
        siteUrl
      );

      const mainListUpdateData = {
        ApproverStatus: "Pending From Approver",
        ApproverComment: removeWhiteSpace(reason),
        SelectedSections: selectedSections,
        RunWF: "Yes",
      };

      await updateDataToSharePoint(MainList, mainListUpdateData, siteUrl, id);

      if (approvalChecks.invoice) {
        const selectedInvoices = invoiceRows.filter(
          (row) => row.invoiceApprovalChecked
        );

        for (const invoice of selectedInvoices) {
          const invoiceUpdateData = {
            PrevInvoiceStatus: invoice.InvoiceStatus || "",
            InvoiceStatus: "Pending Approval",
            RunWF: "Yes",
          };

          await updateDataToSharePoint(
            InvoicelistName,
            invoiceUpdateData,
            siteUrl,
            Number(invoice.itemID || null)
          );
        }
      }

      setModalSnackbar({
        open: true,
        message:
          "Your request to edit has been sent to the Project Manager. Please be patient until it is approved.",
        severity: "success",
      });

      setReason("");
      setApprovalChecks({ client: false, po: false, invoice: false });
      setIsPopupOpen(false);
      showSnackbar(
        "Your request to edit has been sent to the Project Manager. Please be patient until it is approved.",
        "success"
      );

      // Clear form and navigate to dashboard
      setReason("");
      setApprovalChecks({ client: false, po: false, invoice: false });
      setIsPopupOpen(false);

      resetForm();
      await props.refreshCmsDetails();
      setIsLoading(false);
      // alert("Form and data submitted successfully!");
      if (props.onExit) {
        props.onExit();
        return;
      }
      setNavigateToDashboard(true);
      setTimeout(() => {
        setDashboardKey((prev) => prev + 1);
        setNavigateToDashboard(true);
      }, 100);
    } catch (error) {
      console.error("Failed to process edit request:", error);
      setModalSnackbar({
        open: true,
        message: "Something went wrong while sending your edit request.",
        severity: "error",
      });
    } finally {
      setIsSubmitting(false); // Hide loader
    }
  };
  /* const handleSubmitEditRequestApproval = async (
    event?: React.MouseEvent<HTMLButtonElement, MouseEvent>,
    id?: number
  ) => {
    try {
      event?.preventDefault();
      // Ensure id is provided and is a number (narrow type for TS)
      if (typeof id !== "number") {
        console.error("No request id provided to submit edit approval.");
        showSnackbar("No request selected for approval.", "error");
        return;
      }

      if (!reason || reason.trim() === "") {
        showSnackbar("Please provide a reason for the edit request.", "error");
        return;
      }

      if (!approvalChecks || !invoiceRows) {
        console.error("approvalChecks or invoiceRows is undefined.");
        return;
      }

      setIsLoading(true);

      const selectedSections = Object.entries(approvalChecks)
        .filter(([_, checked]) => checked)
        .map(([section, _]) => section)
        .join(", ")
        .trim();

      let selectedInvoiceIDs: any[] = [];
      if (approvalChecks.invoice) {
        selectedInvoiceIDs = invoiceRows
          .filter((row) => row.invoiceApprovalChecked)
          .map((row) => row.itemID)
          .filter((id) => id !== null); // Ensure no null values
      }

      console.log("Selected Sections:", selectedSections);
      console.log("Selected Invoice IDs:", selectedInvoiceIDs);

      // Step 1: Save data in OperationalCMSEditRequest
      const editRequestData = {
        RequestID: id, // now guaranteed to be number
        Reason: reason.trim(),
        UserName: currentUser,
        UserEmail: currentUserEmail,
        SelectedSections: selectedSections,
        Status: "Pending Approval",
        InvoiceID: selectedInvoiceIDs.join(", "),
        ContractID: formData.requestId || "",
      };

      console.log("Edit Request Data:", editRequestData);

      const savedEditRequest = await saveDataToSharePoint(
        OperationalEditRequest,
        editRequestData,
        siteUrl
      );
      console.log("Saved Edit Request:", savedEditRequest);

      // Step 2: Update RunWF in MainList
      const mainListUpdateData = {
        ApproverStatus: "Pending From Approver",
        ApproverComment: removeWhiteSpace(reason),
        SelectedSections: selectedSections,
        RunWF: "Yes",
      };

      console.log("Main List Update Data:", mainListUpdateData);

      const updatedMainList = await updateDataToSharePoint(
        MainList,
        mainListUpdateData,
        siteUrl,
        id // id is a number here
      );
      console.log("Updated MainList:", updatedMainList);

      // Step 3: Update status in InvoicelistName for selected invoices
      if (approvalChecks.invoice) {
        const selectedInvoices = invoiceRows.filter(
          (row) => row.invoiceApprovalChecked
        );

        for (const invoice of selectedInvoices) {
          const invoiceUpdateData = {
            PrevInvoiceStatus: invoice.InvoiceStatus || "",
            InvoiceStatus: "Pending Approval",
            RunWF: "Yes",
          };

          console.log("Invoice Update Data:", invoiceUpdateData);

          await updateDataToSharePoint(
            InvoicelistName,
            invoiceUpdateData,
            siteUrl,
            Number(invoice.itemID || null)
          );
        }
        console.log("Updated Invoice Status for Selected Invoices");
      }

      // Show success message
      showSnackbar(
        "Your request to edit has been sent to the Project Manager. Please be patient until it is approved.",
        "success"
      );

      // Clear form and navigate to dashboard
      setReason("");
      setApprovalChecks({ client: false, po: false, invoice: false });
      setIsPopupOpen(false);

      resetForm();
      await props.refreshCmsDetails();
      setIsLoading(false);
      // alert("Form and data submitted successfully!");
      if (props.onExit) {
        props.onExit();
        return;
      }
      setNavigateToDashboard(true);
      setTimeout(() => {
        setDashboardKey((prev) => prev + 1);
        setNavigateToDashboard(true);
      }, 100);
    } catch (error) {
      console.error("Failed to process edit request:", error);
      console.log(
        "Something went wrong while sending your edit request. Please try again."
      );
    } finally {
      setIsLoading(false);
    }
  };*/

  const handleUpdateEditRequest = async (
    event?: React.MouseEvent<HTMLButtonElement, MouseEvent>
  ) => {
    event?.preventDefault();
    console.log("Update Edit Request clicked");
    console.log(
      "Deleted invoice item IDs (during edit):",
      deletedInvoiceItemIDs
    );

    setIsLoading(true);
    try {
      const selectedSections =
        props.selectedRow?.selectedSections?.toLowerCase();
      if (!selectedSections) {
        alert("No sections selected for update.");
        return;
      }

      // ✅ Create a combined object to update at the end
      const finalMainListData: any = {
        ApproverStatus: "Completed",
        RunWF: "Yes",
        SelectedSections: "",
        ApproverComment: "",
      };

      // ========== CLIENT SECTION ==========
      if (selectedSections.includes("client")) {
        if (!formData.customerEmail.trim()) {
          showSnackbar("Customer Email is required.", "error");
          return;
        }
        if (!formData.location.trim()) {
          showSnackbar("Work Location is required.", "error");
          return;
        }
        if (!formData.workTitle.trim()) {
          showSnackbar("Work Title is required.", "error");
          return;
        }
        if (!formData.workDetail.trim()) {
          showSnackbar("Work Detail is required.", "error");
          return;
        }

        Object.assign(finalMainListData, {
          CustomerEmail: removeWhiteSpace(formData.customerEmail),
          Location: removeWhiteSpace(formData.location),
          WorkTitle: removeWhiteSpace(formData.workTitle),
          WorkDetails: removeWhiteSpace(formData.workDetail),
        });
      }

      // ========== PO SECTION ==========
      if (selectedSections.includes("po")) {
        if (
          formData.poNo &&
          String(formData.poNo).trim() !== "" &&
          !formData.poDate?.trim()
        ) {
          showSnackbar("PO Date is required when PO No is provided.", "error");
          return;
        }
        if (
          formData.poDate &&
          String(formData.poDate).trim() !== "" &&
          !formData.poNo?.trim()
        ) {
          showSnackbar("PO No is required when PO Date is provided.", "error");
          return;
        }

        if (formData.poNo && formData.poNo !== "") {
          const poNoFilterQuery = `$select=PoNo,CompanyName&$filter=CustomerName eq '${encodeURIComponent(
            formData.customerName
          )}' and PoNo eq '${encodeURIComponent(formData.poNo)}' and ID ne ${
            props.selectedRow.id
          }`;
          const poNoData = await getSharePointData(
            { context },
            MainList,
            poNoFilterQuery
          );

          if (poNoData && poNoData.length > 0) {
            showSnackbar(
              "A record with this PO No already exists for the selected Customer Name.",
              "error"
            );
            return;
          }
        }
        // if (!formData.poDate.trim()) {
        //   showSnackbar("PO Date is required.", "error");
        //   return;
        // }

        if (formData.productServiceType?.toLowerCase() !== "azure") {
          if (!formData.poAmount || Number(formData.poAmount) <= 0) {
            showSnackbar(
              "PO Amount is required and must be greater than 0.",
              "error"
            );
            return;
          }
          const poAmt = Number(formData.poAmount) || 0;

          const totalInvoiceAmount = invoiceRows.reduce((sum, r) => {
            if (r.InvoiceStatus === "Credit Note Uploaded") return sum;
            return sum + (Number(r.InvoiceAmount) || 0);
          }, 0);

          const EPS = 0.01;
          if (Math.abs(totalInvoiceAmount - poAmt) > EPS) {
            showSnackbar(
              `Total of invoice amounts (${totalInvoiceAmount.toFixed(
                2
              )}) must equal PO Amount (${poAmt.toFixed(2)}).`,
              "error"
            );
            return;
          }
        }

        Object.assign(finalMainListData, {
          PoNo: formData.poNo,
          PoDate: formData.poDate
            ? moment(formData.poDate, "DD-MM-YYYY").format("YYYY-MM-DD")
            : null,
          POAmount: Number(formData.poAmount),
          BGDate: formData.bgDate
            ? moment(formData.bgDate, "DD-MM-YYYY").format("YYYY-MM-DD")
            : null,
        });
      }

      // ========== INVOICE SECTION ==========
      if (selectedSections.includes("invoice")) {
        const poAmt = Number(formData.poAmount) || 0;

        if (poAmt < 0) {
          showSnackbar("PO Amount cannot be negative.", "error");
          return;
        }

        const totalInvoiceAmount = invoiceRows.reduce((sum, r) => {
          if (r.InvoiceStatus === "Credit Note Uploaded") return sum;
          return sum + (Number(r.InvoiceAmount) || 0);
        }, 0);

        if (totalInvoiceAmount < 0) {
          showSnackbar("Total Invoice Amount cannot be negative.", "error");
          return;
        }
        const EPS = 0.01;
        if (Math.abs(totalInvoiceAmount - poAmt) > EPS) {
          showSnackbar(
            `Total of invoice amounts (${totalInvoiceAmount.toFixed(
              2
            )}) must equal PO Amount (${poAmt.toFixed(2)}).`,
            "error"
          );
          return;
        }

        let isCreditNoteGenerated = false;
        let isCreditNoteUploaded = false;

        for (const row of invoiceRows) {
          if (
            !row.InvoiceDescription ||
            !row.InvoiceAmount ||
            !row.InvoiceDueDate
          ) {
            showSnackbar(
              "One or more invoice rows have missing required fields.",
              "error"
            );
            return;
          }

          const invoiceAmount = Number(row.InvoiceAmount);
          if (invoiceAmount < 0) {
            showSnackbar("Invoice Amount cannot be negative.", "error");
            return;
          }

          const invoiceData: any = {
            Comments: removeWhiteSpace(row.InvoiceDescription),
            PoAmount: Number(row.RemainingPoAmount),
            InvoiceAmount: Number(row.InvoiceAmount),
            InvoiceDueDate: row.InvoiceDueDate
              ? moment(row.InvoiceDueDate, "DD-MM-YYYY").format("YYYY-MM-DD")
              : null,
            EmailBody: row.InvoiceComment,
            RequestID: props.selectedRow.id,
            ClaimNo: row.id,
            PrevInvoiceStatus: "",
            RunWF: "Yes",
          };

          // If row has no itemID, set status to Started
          if (!row.itemID) {
            invoiceData.InvoiceStatus = "Started";
          }

          // If InvoiceStatus is Pending Approval but previous was Generated, revert to previous
          if (
            row.InvoiceStatus === "Pending Approval" &&
            row.PrevInvoiceStatus !== "Generated"
          ) {
            const newStatus = row.PrevInvoiceStatus || " ";
            row.InvoiceStatus = newStatus;
            invoiceData.InvoiceStatus = newStatus;
          }

          // Handle Credit Note status
          if (
            row.InvoiceStatus === "Credit Note Uploaded" &&
            row.CreditNoteStatus === "Uploaded"
          ) {
            invoiceData.CreditNoteStatus = "Completed";
            // invoiceData.PaymentStatus = "No";
            isCreditNoteUploaded = true;
          } else if (
            row.PrevInvoiceStatus === "Generated" &&
            row.CreditNoteStatus === "Pending"
          ) {
            isCreditNoteGenerated = true;
          }

          // Save or update
          if (row.itemID) {
            await updateDataToSharePoint(
              InvoicelistName,
              invoiceData,
              siteUrl,
              row.itemID
            );
          } else {
            await saveDataToSharePoint(InvoicelistName, invoiceData, siteUrl);
          }
        }

        // Handle deletions
        if (
          Array.isArray(deletedInvoiceItemIDs) &&
          deletedInvoiceItemIDs.length
        ) {
          for (const delId of deletedInvoiceItemIDs) {
            await sp.web.lists
              .getByTitle(InvoicelistName)
              .items.getById(Number(delId))
              .delete();
          }
          setDeletedInvoiceItemIDs([]);
        }

        // Add flags at the end
        if (isCreditNoteGenerated) {
          finalMainListData.IsCreditNoteUploaded = "No";
        }
        if (isCreditNoteUploaded) {
          finalMainListData.IsCreditNoteUploaded = "Yes";
        }
      }

      const requestId = props.selectedRow.id;
      const filterQuery = `$filter=RequestID eq ${requestId}&$orderby=ID desc&$top=1`;

      try {
        const operationalData = await getSharePointData(
          props,
          OperationalEditRequest,
          filterQuery
        );
        console.log("Operational Data:", operationalData);

        // Update the CreditNoteUploaded field based on isCreditNoteGenerated
        const updateData = {
          CreditNoteUploaded: finalMainListData.IsCreditNoteUploaded,
        };

        if (operationalData && operationalData.length > 0) {
          const operationalId = operationalData[0].Id;
          await updateDataToSharePoint(
            OperationalEditRequest,
            updateData,
            siteUrl,
            operationalId
          );
        }
      } catch (error) {
        console.error("Error fetching operational data:", error);
      }

      // ✅ Update once at the end
      await updateDataToSharePoint(
        MainList,
        finalMainListData,
        siteUrl,
        props.selectedRow.id
      );
      console.log("Main list updated successfully at the end.");

      showSnackbar("Edit request updated successfully.", "success");

      resetForm();
      await props.refreshCmsDetails();

      if (props.onExit) {
        props.onExit();
        return;
      }

      setNavigateToDashboard(true);
      setTimeout(() => {
        setDashboardKey((prev) => prev + 1);
      }, 100);
    } catch (error) {
      console.error("Error updating edit request:", error);
      showSnackbar("Failed to update edit request.", "error");
    } finally {
      setIsLoading(false);
    }
  };

  const handleReminder = async (
    e: React.MouseEvent<HTMLButtonElement, MouseEvent>,
    id: number
  ) => {
    try {
      e.preventDefault();
    } catch {}

    if (!id) return;

    setIsLoading(true);
    try {
      // 1) Update any Operational Edit Requests for this RequestID that are Pending Approval -> Reminder
      try {
        const opsFilter = `$select=Id,Status&$filter=RequestID eq ${id} and Status eq 'Pending Approval'`;
        const ops = await getSharePointData(
          { context },
          OperationalEditRequest,
          opsFilter
        );
        if (Array.isArray(ops) && ops.length > 0) {
          for (const op of ops) {
            if (op?.Id) {
              await updateDataToSharePoint(
                OperationalEditRequest,
                { Status: "Reminder", ReminderDate: todayDate },
                siteUrl,
                op.Id
              );
            }
          }
        }
      } catch (opErr) {
        console.error(
          "Failed to update Operational Edit Requests to Reminder:",
          opErr
        );
        // continue to update main list even if operational updates fail
      }

      // 2) Update main list ApproverStatus -> Reminder
      const updatedata = {
        ApproverStatus: "Reminder",
        RunWF: "Yes",
      };

      await updateDataToSharePoint(MainList, updatedata, siteUrl, id);

      // 3) UX feedback + refresh
      showSnackbar("Reminder sent to approvers.", "success");
      await fetchOperationalEdits(id);
      // await finalizeAction(false);
      if (typeof props.refreshCmsDetails === "function") {
        await props.refreshCmsDetails();
        await finalizeAction(false);
      }
    } catch (error) {
      console.error("Failed to send reminder:", error);
      showSnackbar("Failed to send reminder. Please try again.", "error");
    } finally {
      setIsLoading(false);
    }
  };

  const handleInvoiceClose = async (e: React.MouseEvent<HTMLButtonElement>) => {
    e.preventDefault();

    // 🔹 Confirmation
    const confirmAction = window.confirm(
      "Are you sure you want to close selected invoices?"
    );

    if (!confirmAction) {
      console.log("User cancelled close operation.");
      return;
    }

    // 🔹 Find selected rows
    const selectedRows = invoiceRows.filter(
      (row) => row.invoiceCloseApprovalChecked === true
    );

    if (selectedRows.length === 0) {
      alert("Please select at least one invoice.");
      return;
    }

    console.log("Selected Rows:", selectedRows);

    const requestData = {
      InvoiceStatus: "Invoice Approval Pending",
      RunWF: "Yes",
    };

    try {
      // 🔹 Loop through selected invoices and update them one by one
      for (const row of selectedRows) {
        if (!row.itemID) {
          console.error("Item ID is missing for row:", row);
          continue;
        }

        await updateDataToSharePoint(
          InvoicelistName,
          requestData,
          props.siteUrl,
          row.itemID
        );

        console.log("Updated Row:", row.itemID);
      }

      // 🔹 Update local UI state
      setInvoiceRows((prevRows) =>
        prevRows.map((r) =>
          r.invoiceCloseApprovalChecked
            ? { ...r, InvoiceStatus: "Invoice Closed" }
            : r
        )
      );

      alert("Selected invoices closed successfully!");
    } catch (error) {
      console.error("Error updating invoice rows:", error);
      alert("Failed to close invoices.");
    }
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "center" }}>
        <Snackbar
          open={snackbar.open}
          autoHideDuration={6000}
          onClose={handleCloseSnackbar}
          anchorOrigin={{ vertical: "top", horizontal: "center" }}
          className="snackbar-container"
          sx={{
            zIndex: 1065, // Higher than the modal's z-index
            opacity: 1,
          }}
        >
          <Alert
            onClose={handleCloseSnackbar}
            severity={
              ["error", "info", "success", "warning"].includes(
                snackbar.severity
              )
                ? (snackbar.severity as
                    | "error"
                    | "info"
                    | "success"
                    | "warning")
                : "info"
            }
          >
            {snackbar.message}
          </Alert>
        </Snackbar>
      </div>
      {isLoading && <LoaderOverlay />}
      <div className="card p-4 shadow">
        <div
          className="headingbar mb-5"
          // style={
          //   props.rowEdit === "Yes" &&
          //   ["Pending From Approver", "Approved", "Hold", "Reminder"].includes(
          //     props.selectedRow?.approverStatus
          //   )
          //     ? {
          //         display: "flex",
          //         justifyContent: "space-between",
          //         alignItems: "center",
          //       }
          //     : {}
          // }
        >
          <h3 className="text-center fw-bold">
            CMS Request
            {/* <button onClick={getVersionHistory}>Get Version History</button> */}
          </h3>

          {/* Conditionally render milestone and note message */}
          {/* {(props.rowEdit === "Yes" &&
            ["Pending From Approver", "Approved", "Hold", "Reminder"].includes(
              props.selectedRow?.approverStatus
            )) 
            ||
            (props.selectedRow?.isCreditNoteUploaded === "No" && (
              <div>
                
                <MilestoneBar
                  status={props.selectedRow?.approverStatus}
                  isCreditNoteUploaded={props.selectedRow?.isCreditNoteUploaded}
                />
              </div>
            ))} */}
          {/* Conditionally render milestone and note message */}
          {(() => {
            const shouldShowMilestone =
              (props.rowEdit === "Yes" &&
                [
                  "Pending From Approver",
                  "Approved",
                  "Hold",
                  "Reminder",
                ].includes(props.selectedRow?.approverStatus)) ||
              props.selectedRow?.isCreditNoteUploaded === "No";

            if (!shouldShowMilestone) return null;

            const creditNoteLabel = getCreditNoteLabel(props.selectedRow);

            return (
              <MilestoneBar
                status={props.selectedRow?.approverStatus}
                creditNoteLabel={creditNoteLabel}
              />
            );
          })()}
        </div>
        <form onSubmit={handleSubmit}>
          <div className="row">
            {/* Requester & Request ID */}
            <div className="col-md-3 mb-3">
              <label className="form-label">
                Requester<span style={{ color: "red" }}>*</span>
              </label>
              <input
                type="text"
                className="form-control"
                name="requester"
                value={formData.requester}
                disabled
                onChange={handleTextFieldChange}
              />
              {errors.requester && (
                <div className="invalid-feedback">{errors.requester}</div>
              )}
            </div>
            <div className="col-md-3 mb-3">
              <label className="form-label">
                Company<span style={{ color: "red" }}>*</span>
              </label>
              <select
                className={`form-select ${
                  errors.companyName ? "is-invalid" : ""
                }`}
                name="companyName"
                value={formData.companyName}
                onChange={handleTextFieldChange}
                disabled={isDisabled}
              >
                <option value="">Select</option>
                {companies.map((option) => (
                  <option key={option} value={option}>
                    {option}
                  </option>
                ))}
              </select>
              {errors.companyName && (
                <div className="invalid-feedback">{errors.companyName}</div>
              )}
            </div>
            <div
              className={`col-md-3 mb-3 ${formData.requestId ? "" : "d-none"}`}
            >
              <label className="form-label">Request ID</label>
              <input
                type="text"
                className={`form-control ${
                  errors.requestId ? "is-invalid" : ""
                }`}
                name="requestId"
                value={formData.requestId}
                onChange={handleTextFieldChange}
                disabled={isDisabled}
              />
              {errors.requestId && (
                <div className="invalid-feedback">{errors.requestId}</div>
              )}
            </div>
          </div>

          <div>
            <div
              className="d-flex align-items-center justify-content-between sectionheader"
              onClick={() => setShowClientWorkDetail((prev) => !prev)}
              aria-expanded={showClientWorkDetail}
              aria-controls="clientWorkDetailCollapse"
            >
              <div style={{ display: "flex", alignItems: "center" }}>
                {/* Approval checkbox (shows only in editable mode and when payment not received) */}
                {props.rowEdit === "Yes" &&
                  requestClosed !== "Yes" &&
                  props.selectedRow?.employeeEmail === currentUserEmail &&
                  ![
                    "Approved",
                    "Hold",
                    "Pending From Approver",
                    "Reminder",
                  ].includes(props.selectedRow?.approverStatus) &&
                  props.selectedRow?.isCreditNoteUploaded !== "No" && (
                    <span
                      className="form-check"
                      onClick={(e) => e.stopPropagation()}
                      style={{ display: "flex", alignItems: "center" }}
                    >
                      <input
                        type="checkbox"
                        id="cbClientWork"
                        className="form-check-input"
                        checked={approvalChecks.client}
                        onChange={(e) =>
                          setApprovalChecks((prev) => ({
                            ...prev,
                            client: e.target.checked,
                          }))
                        }
                      />
                    </span>
                  )}
                <h5 className="mt-3 fw-bold headingColor">
                  Client & Work Detail
                </h5>
              </div>

              <button
                type="button"
                className="btn btn-link"
                onClick={() => setShowClientWorkDetail((prev) => !prev)}
                aria-expanded={showClientWorkDetail}
                aria-controls="clientWorkDetailCollapse"
                style={{ textDecoration: "none", color: "#ffffff" }}
              >
                {showClientWorkDetail ? (
                  <FontAwesomeIcon
                    icon={faAngleUp}
                    onClick={() => setShowClientWorkDetail((prev) => !prev)}
                    aria-expanded={showClientWorkDetail}
                    aria-controls="clientWorkDetailCollapse"
                  />
                ) : (
                  <FontAwesomeIcon
                    icon={faAngleDown}
                    onClick={() => setShowClientWorkDetail((prev) => !prev)}
                    aria-expanded={showClientWorkDetail}
                    aria-controls="clientWorkDetailCollapse"
                  />
                )}
              </button>
            </div>

            <div
              className={`${
                showClientWorkDetail ? "collapse show" : "collapse"
              } sectioncontent`}
              // className={`$collapse show sectioncontent`}
              // id="clientWorkDetailCollapse"
            >
              {" "}
              <div className="row">
                <div className="col-md-3 mb-3">
                  <label className="form-label">
                    Customer Name<span style={{ color: "red" }}>*</span>
                  </label>
                  <select
                    className={`form-select ${
                      errors.customerName ? "is-invalid" : ""
                    }`}
                    name="customerName"
                    value={formData.customerName}
                    onChange={handleTextFieldChange}
                    disabled={isDisabled}
                  >
                    <option value="">Select</option>
                    {customersName.map((option) => (
                      <option key={option} value={option}>
                        {option}
                      </option>
                    ))}
                  </select>
                  {errors.customerName && (
                    <div className="invalid-feedback">
                      {errors.customerName}
                    </div>
                  )}
                </div>
                <div className="col-md-3 mb-3">
                  <label className="form-label">
                    Contract Type<span style={{ color: "red" }}>*</span>
                  </label>
                  <select
                    className={`form-select ${
                      errors.contractType ? "is-invalid" : ""
                    }`}
                    name="contractType"
                    value={formData.contractType}
                    onChange={handleContractTypeChange}
                    disabled={isDisabled}
                  >
                    <option value="">Select</option>
                    {contractTypes.map((option) => (
                      <option key={option} value={option}>
                        {option}
                      </option>
                    ))}
                  </select>{" "}
                  {errors.contractType && (
                    <div className="invalid-feedback">
                      {errors.contractType}
                    </div>
                  )}
                </div>

                <div className="col-md-3 mb-3">
                  <label className="form-label">
                    Product/service Type<span style={{ color: "red" }}>*</span>
                  </label>
                  <select
                    className={`form-select ${
                      errors.productServiceType ? "is-invalid" : ""
                    }`}
                    name="productServiceType"
                    value={formData.productServiceType}
                    onChange={handleTextFieldChange}
                    disabled={isDisabled}
                  >
                    <option value="">Select</option>
                    {productServiceOptions.map((option) => (
                      <option
                        key={option}
                        value={option}
                        // disabled={option.toLowerCase() === "azure" && !formData.customerName.trim()}
                      >
                        {option}
                      </option>
                    ))}
                  </select>
                  {errors.productServiceType && (
                    <div className="invalid-feedback">
                      {errors.productServiceType}
                    </div>
                  )}
                </div>

                <div className="col-md-3 mb-3">
                  <label className="form-label">
                    Govt Contract<span style={{ color: "red" }}>*</span>
                  </label>
                  <div>
                    <div className="form-check form-check-inline">
                      <input
                        className={`form-check-input ${
                          errors.govtContract ? "is-invalid" : ""
                        }`} // Add error class
                        type="radio"
                        name="govtContract"
                        value="Yes"
                        checked={formData.govtContract === "Yes"}
                        onChange={handleTextFieldChange}
                        disabled={isDisabled}
                      />
                      <label className="form-check-label">Yes</label>
                    </div>
                    <div className="form-check form-check-inline">
                      <input
                        className={`form-check-input ${
                          errors.govtContract ? "is-invalid" : ""
                        }`} // Add error class
                        type="radio"
                        name="govtContract"
                        value="No"
                        checked={formData.govtContract === "No"}
                        onChange={handleTextFieldChange}
                        disabled={isDisabled}
                      />
                      <label className="form-check-label">No</label>
                    </div>
                  </div>
                  {errors.govtContract && (
                    <div className="invalid-feedback d-block">
                      {errors.govtContract}
                    </div>
                  )}{" "}
                  {/* Display error */}
                </div>

                <div className="col-md-3 mb-3">
                  <label className="form-label">
                    Customer Email<span style={{ color: "red" }}>*</span>
                  </label>
                  <input
                    type="email"
                    className={`form-control ${
                      errors.customerEmail ? "is-invalid" : ""
                    }`}
                    name="customerEmail"
                    value={formData.customerEmail}
                    onChange={handleTextFieldChange}
                    // disabled={isDisabled }
                    disabled={
                      props.rowEdit === "Yes"
                        ? !(
                            props.selectedRow?.selectedSections
                              ?.toLowerCase()
                              .includes("client") &&
                            props.selectedRow?.approverStatus === "Approved"
                          )
                        : false
                    }
                  />
                  {errors.customerEmail && (
                    <div className="invalid-feedback">
                      {errors.customerEmail}
                    </div>
                  )}
                </div>

                {/* Work Title & Work Detail */}
                <div className="col-md-3 mb-4">
                  <label className="form-label">
                    Account Manager<span style={{ color: "red" }}>*</span>
                  </label>
                  {props.rowEdit === "Yes" ? (
                    <input
                      type="text"
                      className="form-control"
                      value={props.selectedRow.accountMangerTitle || ""}
                      title={props.selectedRow.accountMangerEmail || ""}
                      disabled
                    />
                  ) : (
                    <PeoplePicker
                      context={{
                        msGraphClientFactory:
                          props.context.msGraphClientFactory,
                        spHttpClient: props.context.spHttpClient,
                        absoluteUrl: props.context.pageContext.web.absoluteUrl,
                      }}
                      titleText=""
                      personSelectionLimit={1}
                      showtooltip={true}
                      required={true}
                      onChange={onPeoplePickerChange}
                      principalTypes={[PrincipalType.User]}
                      resolveDelay={1000}
                      disabled={isDisabled}
                      defaultSelectedUsers={activeUsers}
                      styles={{
                        root: {
                          width: "100%",
                          maxHeight: "none",
                          overflow: "hidden",
                        },
                      }}
                    />
                  )}
                  {errors.accountManager && (
                    <div className="invalid-feedback d-block">
                      {errors.accountManager}
                    </div>
                  )}
                </div>
                <div className="col-md-3 mb-4">
                  <label className="form-label">
                    Project Manager<span style={{ color: "red" }}>*</span>
                  </label>
                  {props.rowEdit === "Yes" ? (
                    <input
                      type="text"
                      className="form-control"
                      value={props.selectedRow.projectMangerTitle || ""}
                      title={props.selectedRow.projectMangerEmail || ""}
                      disabled
                    />
                  ) : (
                    <PeoplePicker
                      context={{
                        msGraphClientFactory:
                          props.context.msGraphClientFactory,
                        spHttpClient: props.context.spHttpClient,
                        absoluteUrl: props.context.pageContext.web.absoluteUrl,
                      }}
                      titleText=""
                      personSelectionLimit={1}
                      showtooltip={true}
                      required={true}
                      onChange={onProjectManagerPickerChange}
                      principalTypes={[PrincipalType.User]}
                      resolveDelay={1000}
                      disabled={isDisabled}
                      defaultSelectedUsers={projectManagerUsers}
                      styles={{
                        root: {
                          width: "100%",
                          maxHeight: "none",
                          overflow: "hidden",
                        },
                      }}
                    />
                  )}
                  {errors.projectManager && (
                    <div className="invalid-feedback d-block">
                      {errors.projectManager}
                    </div>
                  )}
                </div>
                <div className="col-md-3 mb-4">
                  <label className="form-label">Project Lead</label>
                  {props.rowEdit === "Yes" ? (
                    <input
                      type="text"
                      className="form-control"
                      value={props.selectedRow.projectLeadTitle || ""}
                      title={props.selectedRow.projectLeadEmail || ""}
                      disabled
                    />
                  ) : (
                    <PeoplePicker
                      context={{
                        msGraphClientFactory:
                          props.context.msGraphClientFactory,
                        spHttpClient: props.context.spHttpClient,
                        absoluteUrl: props.context.pageContext.web.absoluteUrl,
                      }}
                      titleText=""
                      personSelectionLimit={1}
                      showtooltip={true}
                      required={true}
                      onChange={onProjectLeadPickerChange}
                      principalTypes={[PrincipalType.User]}
                      resolveDelay={1000}
                      disabled={isDisabled}
                      defaultSelectedUsers={projectLeadUsers}
                      styles={{
                        root: {
                          width: "100%",
                          maxHeight: "none",
                          overflow: "hidden",
                        },
                      }}
                    />
                  )}
                  {/* {errors.accountManager && (
                    <div className="invalid-feedback d-block">
                      {errors.accountManager}
                    </div>
                  )} */}
                </div>

                {/* Location & Customer Email */}
                <div className="col-md-3 mb-3">
                  <label className="form-label">
                    Work Location<span style={{ color: "red" }}>*</span>
                  </label>
                  <input
                    type="text"
                    className={`form-control ${
                      errors.location ? "is-invalid" : ""
                    }`}
                    name="location"
                    value={formData.location}
                    onChange={handleTextFieldChange}
                    // disabled={isDisabled}
                    disabled={
                      props.rowEdit === "Yes"
                        ? !(
                            props.selectedRow?.selectedSections
                              ?.toLowerCase()
                              .includes("client") &&
                            props.selectedRow?.approverStatus === "Approved"
                          )
                        : false
                    }
                  />

                  {errors.location && (
                    <div className="invalid-feedback">{errors.location}</div>
                  )}
                </div>
                <div className="col-md-3 mb-4">
                  <label className="form-label">
                    Work Title<span style={{ color: "red" }}>*</span>
                  </label>
                  <input
                    type="text"
                    className={`form-control ${
                      errors.workTitle ? "is-invalid" : ""
                    }`}
                    name="workTitle"
                    value={formData.workTitle}
                    onChange={handleTextFieldChange}
                    // disabled={isDisabled}
                    disabled={
                      props.rowEdit === "Yes"
                        ? !(
                            props.selectedRow?.selectedSections
                              ?.toLowerCase()
                              .includes("client") &&
                            props.selectedRow?.approverStatus === "Approved"
                          )
                        : false
                    }
                  />
                  {errors.workTitle && (
                    <div className="invalid-feedback">{errors.workTitle}</div>
                  )}
                </div>
                <div className="col-md-5 mb-7">
                  <label className="form-label">
                    Work Detail<span style={{ color: "red" }}>*</span>
                  </label>
                  <textarea
                    className={`form-control ${
                      errors.workDetail ? "is-invalid" : ""
                    }`}
                    name="workDetail"
                    rows={2}
                    value={formData.workDetail}
                    onChange={handleTextFieldChange}
                    // disabled={isDisabled}
                    disabled={
                      props.rowEdit === "Yes"
                        ? !(
                            props.selectedRow?.selectedSections
                              ?.toLowerCase()
                              .includes("client") &&
                            props.selectedRow?.approverStatus === "Approved"
                          )
                        : false
                    }
                  />
                  {errors.workDetail && (
                    <div className="invalid-feedback">{errors.workDetail}</div>
                  )}
                </div>
              </div>
            </div>
          </div>

          <div className="mt-4">
            <div
              className="d-flex align-items-center justify-content-between sectionheader"
              onClick={() => setShowPODetails((prev) => !prev)}
              aria-expanded={showPODetails}
              aria-controls="poDetailsCollapse"
            >
              <div style={{ display: "flex", alignItems: "center" }}>
                {" "}
                {/* PO Details */}
                {props.rowEdit === "Yes" &&
                  requestClosed !== "Yes" &&
                  props.selectedRow?.employeeEmail === currentUserEmail &&
                  ![
                    "Approved",
                    "Hold",
                    "Pending From Approver",
                    "Reminder",
                  ].includes(props.selectedRow.approverStatus) &&
                  props.selectedRow?.isCreditNoteUploaded !== "No" && (
                    <span
                      className="form-check "
                      onClick={(e) => e.stopPropagation()}
                      style={{ display: "flex", alignItems: "center" }}
                    >
                      <input
                        type="checkbox"
                        id="cbPoDetails"
                        className="form-check-input"
                        checked={approvalChecks.po}
                        onChange={(e) =>
                          setApprovalChecks((prev) => ({
                            ...prev,
                            po: e.target.checked,
                          }))
                        }
                      />
                    </span>
                  )}
                <h5 className="mt-2 fw-bold headingColor">PO Details</h5>
              </div>
              <button
                type="button"
                className="btn btn-link"
                onClick={() => setShowPODetails((prev) => !prev)}
                aria-expanded={showPODetails}
                aria-controls="poDetailsCollapse"
                style={{ textDecoration: "none", color: "#ffffff" }}
              >
                {showPODetails ? (
                  <FontAwesomeIcon
                    icon={faAngleUp}
                    onClick={() => setShowPODetails((prev) => !prev)}
                    aria-expanded={showPODetails}
                    aria-controls="poDetailsCollapse"
                  />
                ) : (
                  <FontAwesomeIcon
                    icon={faAngleDown}
                    onClick={() => setShowPODetails((prev) => !prev)}
                    aria-expanded={showPODetails}
                    aria-controls="poDetailsCollapse"
                  />
                )}
              </button>
            </div>

            <div
              className={`${
                showPODetails ? "collapse show" : "collapse"
              } sectioncontent`}
              id="poDetailsCollapse"
            >
              <div className="row">
                {" "}
                <div className="col-md-3 mb-3">
                  <label className="form-label">
                    PO No{" "}
                    {formData.bgRequired === "Yes" && (
                      <span style={{ color: "red" }}>*</span>
                    )}
                  </label>
                  <input
                    type="text"
                    className={`form-control ${
                      errors.poNo ? "is-invalid" : ""
                    }`}
                    name="poNo"
                    value={formData.poNo}
                    onChange={handleTextFieldChange}
                    // disabled={
                    //   requestClosed === "Yes" ||
                    //   (props.selectedRow?.employeeEmail !==
                    //     props.context.pageContext.user.email &&
                    //     props.rowEdit === "Yes")
                    // }

                    // disabled={
                    //   (isDisabled &&
                    //     props.selectedRow.poNo &&
                    //     props.rowEdit === "Yes") ||
                    //   (props.selectedRow?.employeeEmail !==
                    //     props.context.pageContext.user.email &&
                    //     props.rowEdit === "Yes")
                    // }
                    disabled={
                      props.rowEdit === "Yes"
                        ? !(
                            props.selectedRow?.employeeEmail ===
                              currentUserEmail &&
                            props.selectedRow?.selectedSections
                              ?.toLowerCase()
                              .includes("po") &&
                            props.selectedRow?.approverStatus === "Approved"
                          )
                        : false
                    }
                  />
                  {errors.poNo && (
                    <div className="invalid-feedback">{errors.poNo}</div>
                  )}
                </div>
                {/* <div className="col-md-3 mb-3">
              <label className="form-label">PO Date</label>
              <input
                type="date"
               
                className={`form-control ${errors.poDate ? "is-invalid" : ""}`}
                name="poDate"
                value={formData.poDate}
                onChange={handleTextFieldChange}
                disabled={isDisabled}
              />
              {errors.poDate && (
                <div className="invalid-feedback">{errors.poDate}</div>
              )}
            </div> */}
                <div className="col-md-3 mb-3">
                  <label className="form-label">
                    PO Date
                    {(formData.bgRequired === "Yes" ||
                      formData.poNo.trim() !== "") && (
                      <span style={{ color: "red" }}>*</span>
                    )}
                  </label>

                  <DatePicker
                    className={`form-control ${
                      errors.poDate ? "is-invalid" : ""
                    }`}
                    name="poDate"
                    format="DD-MM-YYYY"
                    value={
                      formData.poDate &&
                      moment(formData.poDate, "DD-MM-YYYY", true).isValid()
                        ? moment(formData.poDate, "DD-MM-YYYY")
                        : null
                    }
                    onChange={(date) => {
                      let updatedPoDate = date ? date.format("DD-MM-YYYY") : "";
                      let updatedFormData = {
                        ...formData,
                        poDate: updatedPoDate,
                      };
                      // Auto-set bgDate if BG Required is Yes
                      if (formData.bgRequired === "Yes" && date) {
                        updatedFormData.bgDate = date
                          .clone()
                          .add(20, "days")
                          .format("DD-MM-YYYY");
                      }
                      setFormData(updatedFormData);
                    }}
                    // disabled={
                    //   (isDisabled &&
                    //     props.selectedRow.poDate &&
                    //     props.rowEdit === "Yes") ||
                    //   (props.selectedRow?.employeeEmail !==
                    //     props.context.pageContext.user.email &&
                    //     props.rowEdit === "Yes")
                    // }
                    disabled={
                      props.rowEdit === "Yes"
                        ? !(
                            props.selectedRow?.employeeEmail ===
                              currentUserEmail &&
                            props.selectedRow?.selectedSections
                              ?.toLowerCase()
                              .includes("po") &&
                            props.selectedRow?.approverStatus === "Approved"
                          )
                        : false
                    }
                    // disabled={
                    //   requestClosed === "Yes" ||
                    //   (props.selectedRow?.employeeEmail !==
                    //     props.context.pageContext.user.email &&
                    //     props.rowEdit === "Yes")
                    // }

                    // Disable if PO Date is already set
                    disabledDate={(current) => {
                      return current && current > moment().endOf("day");
                    }}
                  />
                  {errors.poDate && (
                    <div className="invalid-feedback">{errors.poDate}</div>
                  )}
                </div>
                <div className="col-md-3 mb-3">
                  <label className="form-label">
                    PO Amount{" "}
                    {formData.productServiceType?.toLowerCase() !== "azure" && (
                      <span style={{ color: "red" }}>*</span>
                    )}
                  </label>

                  <input
                    type="number"
                    className={`form-control ${
                      errors.poAmount ? "is-invalid" : ""
                    }`}
                    name="poAmount"
                    step="any"
                    value={formData.poAmount}
                    min={1}
                    onChange={(e) => {
                      const value = e.target.value;
                      if (props.rowEdit === "Yes") {
                        if (Number(value) < 0) return;
                        // Update PO Amount and reset invoiceCriteria
                        setFormData((prev) => ({
                          ...prev,
                          poAmount: value,
                        }));
                      } else {
                        if (Number(value) < 0) return;
                        // Update PO Amount and reset invoiceCriteria
                        setFormData((prev) => ({
                          ...prev,
                          poAmount: value,
                          invoiceCriteria: "",
                        }));
                        // Reset invoice rows
                        setInvoiceRows([
                          {
                            id: 1,
                            InvoiceDescription: "",
                            RemainingPoAmount: "",
                            InvoiceAmount: "",
                            InvoiceDueDate: "",
                            InvoiceProceedDate: "",
                            InvoiceComment: "",
                            showProceed: false,
                            InvoiceStatus: "",
                            userInGroup: false,
                            employeeEmail: "",
                            itemID: null,
                            InvoiceNo: "",
                            InvoiceDate: "",
                            InvoiceTaxAmount: "",
                            ClaimNo: null,
                            PrevInvoiceStatus: "",
                            CreditNoteStatus: "",
                            RequestID: "",
                            DocId: "",
                            PendingAmount: "",
                            InvoiceFileID: "",
                            invoiceApprovalChecked: false, // Initialize here
                            invoiceCloseApprovalChecked: false, // Initialize here
                          },
                        ]);
                      }
                    }}
                    // disabled={isDisabled}
                    disabled={
                      props.rowEdit === "Yes"
                        ? !(
                            props.selectedRow?.employeeEmail ===
                              currentUserEmail &&
                            props.selectedRow?.selectedSections
                              ?.toLowerCase()
                              .includes("po") &&
                            props.selectedRow?.approverStatus === "Approved"
                          )
                        : false
                    }
                  />

                  {errors.poAmount && (
                    <div className="invalid-feedback">{errors.poAmount}</div>
                  )}
                </div>
                <div className="col-md-3 mb-3">
                  <label className="form-label">
                    Currency<span style={{ color: "red" }}>*</span>
                  </label>
                  <select
                    className={`form-select ${
                      errors.currency ? "is-invalid" : ""
                    }`}
                    name="currency"
                    // value={formData.currency}
                    value={formData.currency}
                    onChange={handleTextFieldChange}
                    disabled={isDisabled}
                  >
                    <option value="">Select</option>
                    {currencyName.map((option) => (
                      <option key={option} value={option}>
                        {option}
                      </option>
                    ))}
                  </select>
                  {errors.currency && (
                    <div className="invalid-feedback">{errors.currency}</div>
                  )}
                </div>
                {formData.productServiceType?.toLowerCase() === "resource" && (
                  <>
                    {console.log(
                      "Rendering fields for Product/service Type: Resource"
                    )}{" "}
                    {/* Debugging */}
                    <div className="col-md-3 mb-3">
                      <label className="form-label">Start Date</label>
                      <span style={{ color: "red" }}>*</span>
                      <DatePicker
                        format="DD-MM-YYYY"
                        className={`form-control ${
                          errors.startDate ? "is-invalid" : ""
                        }`}
                        value={
                          formData.startDate
                            ? moment(formData.startDate, "DD-MM-YYYY")
                            : null
                        }
                        onChange={(date) => {
                          const formatted = date
                            ? date.format("DD-MM-YYYY")
                            : "";
                          setFormData((prev) => ({
                            ...prev,
                            startDate: formatted,
                            endDate: "",
                            invoiceCriteria: "", // Reset Invoice Criteria dropdown
                          }));

                          setInvoiceRows([
                            {
                              id: 1,
                              InvoiceDescription: "",
                              RemainingPoAmount: "",
                              InvoiceAmount: "",
                              InvoiceDueDate: "",
                              InvoiceProceedDate: "",
                              InvoiceComment: "",
                              showProceed: false,
                              InvoiceStatus: "",
                              userInGroup: false,
                              employeeEmail: "",
                              itemID: null,
                              InvoiceNo: "",
                              InvoiceDate: "",
                              InvoiceTaxAmount: "",
                              ClaimNo: null,
                              PrevInvoiceStatus: "",
                              CreditNoteStatus: "",
                              RequestID: "",
                              DocId: "",
                              PendingAmount: "",
                              InvoiceFileID: "",
                              invoiceApprovalChecked: false, // Initialize here
                              invoiceCloseApprovalChecked: false, // Initialize here
                            },
                          ]);
                        }}
                        disabled={isDisabled}
                      />
                      {errors.startDate && (
                        <div className="invalid-feedback">
                          {errors.startDate}
                        </div>
                      )}
                    </div>
                    <div className="col-md-3 mb-3">
                      <label className="form-label">End Date</label>
                      <span style={{ color: "red" }}>*</span>

                      <DatePicker
                        format="DD-MM-YYYY"
                        className={`form-control ${
                          errors.endDate ? "is-invalid" : ""
                        }`}
                        value={
                          formData.endDate
                            ? moment(formData.endDate, "DD-MM-YYYY")
                            : null
                        }
                        onChange={(date) => {
                          if (!formData.startDate) {
                            alert("Please select a start date first.");
                            return;
                          }
                          const formatted = date
                            ? date.format("DD-MM-YYYY")
                            : "";
                          setFormData((prev) => ({
                            ...prev,
                            endDate: formatted,

                            invoiceCriteria: "",
                          }));
                          setInvoiceRows([
                            {
                              id: 1,
                              InvoiceDescription: "",
                              RemainingPoAmount: "",
                              InvoiceAmount: "",
                              InvoiceDueDate: "",
                              InvoiceProceedDate: "",
                              InvoiceComment: "",
                              showProceed: false,
                              InvoiceStatus: "",
                              userInGroup: false,
                              employeeEmail: "",
                              itemID: null,
                              InvoiceNo: "",
                              InvoiceDate: "",
                              InvoiceTaxAmount: "",
                              ClaimNo: null,
                              PrevInvoiceStatus: "",
                              CreditNoteStatus: "",
                              RequestID: "",
                              DocId: "",
                              PendingAmount: "",
                              InvoiceFileID: "",
                              invoiceApprovalChecked: false, // Initialize here
                              invoiceCloseApprovalChecked: false, // Initialize here
                            },
                          ]);
                        }}
                        disabledDate={(current) =>
                          formData.startDate
                            ? current &&
                              (current.isBefore(
                                moment(
                                  formData.startDate,
                                  "DD-MM-YYYY"
                                ).toDate(),
                                "day"
                              ) ||
                                current.isSame(
                                  moment(
                                    formData.startDate,
                                    "DD-MM-YYYY"
                                  ).toDate(),
                                  "day"
                                ))
                            : false
                        }
                        disabled={isDisabled}
                      />

                      {errors.endDate && (
                        <div className="invalid-feedback">{errors.endDate}</div>
                      )}
                    </div>
                    <div className="col-md-3 mb-3">
                      <label className="form-label">Payment Mode</label>
                      <span style={{ color: "red" }}>*</span>

                      {/* <select
                        name="paymentMode"
                        className={`form-select ${errors.paymentMode ? "is-invalid" : ""
                          }`}
                        value={formData.paymentMode} // Bind the value to formData.invoiceCriteria
                        onChange={(e) => {
                         // Proceed with the change if PO Amount is filled
                          setFormData((prev) => ({
                              ...prev,
                              invoiceCriteria: "", 
                            }));

                       handleInvoiceCriteriaChange(e);
                        }}
                        disabled={isDisabled}
                      >
                        <option value="">Select</option>
                        <option value="Pre">Pre</option>
                        <option value="Post">Post</option>
                       
                      
                      </select> */}
                      <select
                        name="paymentMode"
                        className={`form-select ${
                          errors.paymentMode ? "is-invalid" : ""
                        }`}
                        value={formData.paymentMode}
                        onChange={(e) => {
                          setFormData((prev) => ({
                            ...prev,
                            paymentMode: e.target.value,
                            invoiceCriteria: "", // Reset invoiceCriteria when paymentMode changes
                          }));
                          setInvoiceRows([
                            {
                              id: 1,
                              InvoiceDescription: "",
                              RemainingPoAmount: "",
                              InvoiceAmount: "",
                              InvoiceDueDate: "",
                              InvoiceProceedDate: "",
                              InvoiceComment: "",
                              showProceed: false,
                              InvoiceStatus: "",
                              userInGroup: false,
                              employeeEmail: "",
                              itemID: null,
                              InvoiceNo: "",
                              InvoiceDate: "",
                              InvoiceTaxAmount: "",
                              ClaimNo: null,
                              PrevInvoiceStatus: "",
                              CreditNoteStatus: "",
                              RequestID: "",
                              DocId: "",
                              PendingAmount: "",
                              InvoiceFileID: "",
                              invoiceApprovalChecked: false, // Initialize here
                              invoiceCloseApprovalChecked: false, // Initialize here
                            },
                          ]);
                          // Do NOT call handleInvoiceCriteriaChange(e) here
                        }}
                        disabled={isDisabled}
                      >
                        <option value="">Select</option>
                        <option value="Pre">Pre</option>
                        <option value="Post">Post</option>
                      </select>
                      {errors.paymentMode && (
                        <div className="invalid-feedback">
                          {errors.paymentMode}
                        </div>
                      )}
                    </div>
                    <div className="col-md-3 mb-3">
                      <label className="form-label">Invoice Criteria</label>
                      <span style={{ color: "red" }}>*</span>

                      <select
                        name="invoiceCriteria"
                        className={`form-select ${
                          errors.invoiceCriteria ? "is-invalid" : ""
                        }`}
                        value={formData.invoiceCriteria} // Bind the value to formData.invoiceCriteria
                        onChange={(e) => {
                          if (!formData.poAmount.trim()) {
                            alert(
                              "Please fill PO Amount before selecting Invoice Criteria."
                            );
                            setFormData((prev) => ({
                              ...prev,
                              invoiceCriteria: "", // Reset the dropdown value
                            }));
                            return;
                          } else if (!formData.endDate) {
                            alert(
                              "Please select Start Date & End Date before selecting Invoice Criteria."
                            );
                            setFormData((prev) => ({
                              ...prev,
                              invoiceCriteria: "", // Reset the dropdown value
                            }));
                            return;
                          } else if (!formData.paymentMode) {
                            alert(
                              "Please select Payment Mode before selecting Invoice Criteria."
                            );
                            setFormData((prev) => ({
                              ...prev,
                              invoiceCriteria: "", // Reset the dropdown value
                            }));
                            return;
                          }
                          handleInvoiceCriteriaChange(e); // Proceed with the change if PO Amount is filled
                        }}
                        disabled={isDisabled}
                      >
                        <option value="">Select</option>
                        <option value="Yearly">Yearly</option>
                        <option value="Half-yearly">Half-yearly</option>
                        <option value="Quarterly">Quarterly</option>
                        <option value="2-Monthly">2-Monthly</option>
                        <option value="Monthly">Monthly</option>
                      </select>
                      {errors.invoiceCriteria && (
                        <div className="invalid-feedback">
                          {errors.invoiceCriteria}
                        </div>
                      )}
                    </div>
                  </>
                )}
                <div className="col-md-2 mb-2">
                  <label className="form-label">
                    BG Required<span style={{ color: "red" }}>*</span>
                  </label>
                  <div>
                    <div className="form-check form-check-inline">
                      <input
                        className={`form-check-input ${
                          errors.bgRequired ? "is-invalid" : ""
                        }`} // Add error class
                        type="radio"
                        name="bgRequired"
                        value="Yes"
                        checked={formData.bgRequired === "Yes"}
                        onChange={handleTextFieldChange}
                        disabled={isDisabled}
                      />
                      <label className="form-check-label">Yes</label>
                    </div>
                    <div className="form-check form-check-inline">
                      <input
                        className={`form-check-input ${
                          errors.bgRequired ? "is-invalid" : ""
                        }`} // Add error class
                        type="radio"
                        name="bgRequired"
                        value="No"
                        checked={formData.bgRequired === "No"}
                        onChange={handleTextFieldChange}
                        disabled={isDisabled}
                      />
                      <label className="form-check-label">No</label>
                    </div>
                  </div>
                  {errors.bgRequired && (
                    <div className="invalid-feedback d-block">
                      {errors.bgRequired}
                    </div>
                  )}{" "}
                  {/* Display error */}
                </div>
                {formData.bgRequired === "Yes" && (
                  <div className="col-md-3 mb-3">
                    <label className="form-label">
                      BG Required By<span style={{ color: "red" }}>*</span>
                    </label>
                    <DatePicker
                      className={`form-control ${
                        errors.bgDate ? "is-invalid" : ""
                      }`}
                      name="bgDate"
                      format="DD-MM-YYYY"
                      value={
                        formData.bgDate
                          ? moment(formData.bgDate, "DD-MM-YYYY")
                          : null
                      }
                      onChange={(date) => {
                        setFormData((prev) => ({
                          ...prev,
                          bgDate: date ? date.format("DD-MM-YYYY") : "",
                        }));
                      }}
                      // disabled={isDisabled}
                      disabled
                    />
                    {errors.bgDate && (
                      <div className="invalid-feedback">{errors.bgDate}</div>
                    )}
                  </div>
                )}
                {/* </div>
          <div className="row"> */}
                {/* <h5 className="mt-3 fw-bold mb-4 text-decoration-underline headingColor">
              PO Attachment
            </h5> */}
                {poUploadedAttachmentFiles.length <= 0 ? (
                  <>
                    <div className="col-md-3 mb-7">
                      <label className="form-label">Comment</label>
                      <textarea
                        className={`form-control ${
                          errors.poComment ? "is-invalid" : ""
                        }`}
                        name="poComment"
                        rows={2}
                        value={formData.poComment}
                        // disabled={poUploadedAttachmentFiles.length > 0 || (isDisabled && props.selectedRow.poNo && props.rowEdit === "Yes") || (props.selectedRow?.employeeEmail !== props.context.pageContext.user.email)}

                        // disabled={
                        // poUploadedAttachmentUploadedFiles.length > 0 ||
                        disabled={
                          poUploadedAttachmentFiles.length > 0 ||
                          requestClosed === "Yes" ||
                          (props.selectedRow?.employeeEmail !==
                            props.context.pageContext.user.email &&
                            props.rowEdit === "Yes")
                        }
                        onChange={handleTextFieldChange}
                      />
                      {errors.poComment && (
                        <div className="invalid-feedback">
                          {errors.poComment}
                        </div>
                      )}
                    </div>
                    <div className="col-md-4 mb-3">
                      <label className="form-label">
                        Upload Attachment
                        {(formData.poNo.trim() ||
                          formData.bgRequired === "Yes") && (
                          <span style={{ color: "red" }}>*</span>
                        )}
                      </label>
                      <div className="input-group">
                        <input
                          type="file"
                          className="form-control"
                          name="pofile"
                          id="poFileInput"
                          ref={poFileInputRef}
                          onChange={handlePoFileChange}
                          // disabled={
                          //   poUploadedAttachmentFiles.length > 0 || isUserInAdminM || (isDisabled && props.selectedRow.poNo && props.rowEdit === "Yes") || (props.selectedRow?.employeeEmail !== props.context.pageContext.user.email)
                          // }
                          disabled={
                            poUploadedAttachmentFiles.length > 0 ||
                            isUserInAdminM ||
                            requestClosed === "Yes" ||
                            (props.selectedRow?.employeeEmail !==
                              props.context.pageContext.user.email &&
                              props.rowEdit === "Yes")
                          }
                        />
                        <button
                          type="button"
                          className="btn btn-success"
                          onClick={uploadPO}
                          // disabled={
                          //   poUploadedAttachmentFiles.length > 0 ||
                          //   isUserInAdminM ||
                          //   (isDisabled && props.selectedRow.poNo && props.rowEdit === "Yes") ||
                          //   (props.selectedRow?.employeeEmail !== props.context.pageContext.user.email && props.rowEdit === "Yes")
                          // }
                          disabled={
                            poUploadedAttachmentFiles.length > 0 ||
                            isUserInAdminM ||
                            requestClosed === "Yes" ||
                            (props.selectedRow?.employeeEmail !==
                              props.context.pageContext.user.email &&
                              props.rowEdit === "Yes")
                          }
                        >
                          {/* Upload */}
                          <FontAwesomeIcon icon={faUpload} />
                        </button>
                      </div>
                    </div>
                  </>
                ) : null}
              </div>

              <div className="row mt-4">
                <div className="col-12">
                  {poUploadedAttachmentFiles.length > 0 ? (
                    <div className="card">
                      {/* <div className="card-header bg-primary text-white">
                    <h5 className="mb-0">PO Attached Files</h5>
                  </div> */}
                      <div className="card-body p-0">
                        <table className="table table-striped table-bordered mb-0">
                          <thead className="table-light">
                            <tr>
                              <th scope="col">S.No</th>
                              <th scope="col">Document Name</th>{" "}
                              {/* FileLeafRef */}
                              <th scope="col">Attachment Type</th>
                              <th scope="col">Comment</th>
                              <th scope="col">Action</th>
                            </tr>
                          </thead>
                          <tbody>
                            {poUploadedAttachmentFiles
                              .filter((file) => file.AttachmentType === "PO")
                              .map((file, index) => (
                                <tr key={file.Id}>
                                  <td>{index + 1}</td>
                                  <td>{file.FileLeafRef}</td>
                                  <td>{file.AttachmentType}</td>
                                  <td>{file.Comment}</td>
                                  <td>
                                    <button
                                      type="button"
                                      className="btn btn-sm btn-outline-primary me-2"
                                      onClick={(e) => {
                                        const viewUrl = getViewUrl(file);
                                        handlePoView(
                                          e,
                                          viewUrl,
                                          file.FileLeafRef
                                        );
                                      }}
                                    >
                                      <FontAwesomeIcon icon={faEye} />{" "}
                                    </button>
                                    <button
                                      type="button"
                                      className="btn btn-sm btn-outline-success me-2"
                                      onClick={(e) =>
                                        handlePoDownload(e, file.EncodedAbsUrl)
                                      }
                                    >
                                      <FontAwesomeIcon icon={faFileArrowDown} />{" "}
                                    </button>

                                    {/* {props.selectedRow?.UserEmail === currentUserEmail && ( */}
                                    {file.UserEmail === currentUserEmail &&
                                      requestClosed !== "Yes" && (
                                        <button
                                          type="button"
                                          className="btn btn-sm btn-outline-danger"
                                          onClick={() => handlePoDelete(file)}
                                          disabled={
                                            isUserInAdminM ||
                                            requestClosed === "Yes"
                                          }
                                        >
                                          <FontAwesomeIcon icon={faTrash} />
                                        </button>
                                      )}
                                    {/* )} */}
                                  </td>
                                </tr>
                              ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  ) : (
                    // <div className="alert alert-info mt-2">No Attachment</div>
                    (requestClosed === "Yes" || isUserInAdminM) && (
                      <div className="alert alert-info mt-2">No Attachment</div>
                    )
                  )}

                  {/* } */}
                </div>
              </div>
            </div>
          </div>

          {props.rowEdit === "Yes" && (
            <div className="mt-4">
              <div className="mt-4">
                <div
                  className="d-flex align-items-center justify-content-between sectionheader"
                  onClick={() => setShowBGDetails((prev) => !prev)}
                  aria-expanded={showBGDetails}
                  aria-controls="bgDetailsCollapse"
                >
                  <h5 className="m-2 fw-bold headingColor">BG Details</h5>
                  <button
                    type="button"
                    className="btn btn-link"
                    onClick={() => setShowBGDetails((prev) => !prev)}
                    aria-expanded={showBGDetails}
                    aria-controls="bgDetailsCollapse"
                    style={{ textDecoration: "none", color: "#ffffff" }}
                  >
                    {showBGDetails ? (
                      <FontAwesomeIcon
                        icon={faAngleUp}
                        onClick={() => setShowBGDetails((prev) => !prev)}
                        aria-expanded={showBGDetails}
                        aria-controls="bgDetailsCollapse"
                      />
                    ) : (
                      <FontAwesomeIcon
                        icon={faAngleDown}
                        onClick={() => setShowBGDetails((prev) => !prev)}
                        aria-expanded={showBGDetails}
                        aria-controls="bgDetailsCollapse"
                      />
                    )}
                  </button>
                </div>
                <div
                  className={`${
                    showBGDetails ? "collapse show" : "collapse"
                  } sectioncontent`}
                  id="bgDetailsCollapse"
                >
                  {/* BG Details Section - Only show BG End Date and Upload if NOT admin */}
                  {requestClosed !== "Yes" &&
                  !isUserInAdminM &&
                  props.rowEdit === "Yes" &&
                  (props.selectedRow?.employeeEmail ===
                    props.context.pageContext.user.email ||
                    isUserInGroupM) ? (
                    <div className="row mb-3">
                      <div className="col-md-3">
                        <label htmlFor="bgEndDate" className="form-label">
                          BG End Date<span style={{ color: "red" }}>*</span>
                        </label>

                        <DatePicker
                          className={`form-control ${
                            errors.bgEndDate ? "is-invalid" : ""
                          }`}
                          name="bgEndDate"
                          format="DD-MM-YYYY"
                          value={
                            formData.bgEndDate
                              ? moment(formData.bgEndDate, "DD-MM-YYYY")
                              : null
                          }
                          onChange={(date) => {
                            setFormData((prev) => ({
                              ...prev,
                              bgEndDate: date ? date.format("DD-MM-YYYY") : "",
                            }));
                          }}
                          // disabled={isDisabled}
                          disabled={isUserInAdminM || requestClosed === "Yes"}
                        />
                      </div>
                      <div className="col-md-6">
                        <label htmlFor="fileInput" className="form-label">
                          Upload File<span style={{ color: "red" }}>*</span>
                        </label>
                        <div className="input-group">
                          <input
                            type="file"
                            className="form-control"
                            id="fileInput"
                            ref={bgFileInputRef}
                            onChange={UploadBgFileChange}
                            disabled={isUserInAdminM || requestClosed === "Yes"}
                            // disabled={isUserInAdminM  || (props.selectedRow && props.selectedRow.isPaymentReceived === "Yes")}
                          />
                          <button
                            className="btn btn-success"
                            onClick={UploadBgFile}
                            disabled={isUserInAdminM || requestClosed === "Yes"}
                          >
                            {/* Upload */}
                            <FontAwesomeIcon icon={faUpload} />
                          </button>
                        </div>
                      </div>
                    </div>
                  ) : null}
                  <div className="row mt-4">
                    {uploadedBGFiles.length > 0 ? (
                      <div className="card-body p-0">
                        <table className="table table-striped table-bordered mb-0">
                          <thead className="table-light">
                            <tr>
                              <th scope="col">S.No</th>
                              {/* <th>FileID</th>
                                                    <th>DocID</th> */}
                              <th scope="col">Document Name</th>
                              <th scope="col">BG Relased Date</th>
                              <th scope="col">Actions</th>
                            </tr>
                          </thead>
                          <tbody>
                            {uploadedBGFiles.map((file, index) => (
                              <tr key={file.Id}>
                                <td>{index + 1}</td>
                                {/* <td>{file.FileID}</td>
                                                            <td>{file.DocID}</td> */}

                                <td>{file.FileLeafRef}</td>
                                <td>
                                  {console.log(
                                    "BGReleaseDate:",
                                    file.BGDate,
                                    "Formatted:",
                                    formatDateForTable(file.BGDate)
                                  )}
                                  {formatDateForTable(file.BGDate)}
                                </td>
                                <td>
                                  {(() => {
                                    return (
                                      <button
                                        className="btn btn-sm btn-outline-primary"
                                        onClick={(e) => {
                                          const viewUrl = getViewUrl(file);
                                          handleViewBGFile(
                                            e,
                                            viewUrl,
                                            file.FileLeafRef
                                          );
                                        }}
                                      >
                                        {/* View */}
                                        <FontAwesomeIcon icon={faEye} />
                                      </button>
                                    );
                                  })()}

                                  <button
                                    className="btn btn-sm btn-outline-success ms-2"
                                    onClick={(e) =>
                                      handleBgDownload(e, file.EncodedAbsUrl)
                                    }
                                  >
                                    <FontAwesomeIcon icon={faFileArrowDown} />
                                  </button>

                                  <button
                                    className="btn btn-sm btn-outline-warning ms-2"
                                    onClick={(e) => handleEditClick(e, file)}
                                    // disabled={
                                    //   isUserInAdminM || requestClosed === "Yes"
                                    // }
                                    // disabled={
                                    //   requestClosed === "Yes"
                                    // }
                                  >
                                    {" "}
                                    {/* Edit */}
                                    <FontAwesomeIcon icon={faPenToSquare} />
                                  </button>
                                  {file.UserEmail === currentUserEmail && (
                                    <button
                                      className="btn btn-sm btn-outline-danger ms-2"
                                      onClick={(e) =>
                                        handleDeleteBGFile(e, file)
                                      }
                                      disabled={
                                        isUserInAdminM ||
                                        requestClosed === "Yes"
                                      }
                                    >
                                      {/* Delete */}
                                      <FontAwesomeIcon icon={faTrash} />
                                    </button>
                                  )}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    ) : (
                      (requestClosed !== "No" || isUserInAdminM) && (
                        <div className="alert alert-info mt-2">
                          No Attachment
                        </div>
                      )
                    )}
                  </div>
                </div>
              </div>

              {/* Edit Modal */}
              <Modal
                show={showEditBGModal}
                onHide={handleCloseBGModal}
                centered
                dialogClassName="custommodalwidth"
              >
                <Modal.Header closeButton>
                  <Modal.Title>Edit BG Details</Modal.Title>
                </Modal.Header>
                <Modal.Body>
                  {isLoading && (
                    <div
                      style={{
                        position: "absolute",
                        top: 0,
                        left: 0,
                        right: 0,
                        bottom: 0,
                        background: "rgba(255,255,255,0.7)",
                        zIndex: 200000, // higher than modal content
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                      }}
                    >
                      <Spinner animation="border" variant="primary" />
                      <span className="ms-3">Processing...</span>
                    </div>
                  )}

                  <div className="mb-3">
                    <label className="form-label">Is BG Release</label>
                    <div>
                      <div className="form-check form-check-inline">
                        <input
                          className="form-check-input"
                          type="radio"
                          name="isBGRelease"
                          value="No"
                          id="isBGReleaseNo"
                          checked={isBGRelease === "No"}
                          onChange={handleBGReleaseChange}
                          disabled={
                            RelasedBGFiles.length > 0 ||
                            requestClosed === "Yes" ||
                            isUserInAdminM
                          }
                        />
                        <label
                          className="form-check-label"
                          htmlFor="isBGReleaseNo"
                        >
                          No
                        </label>
                      </div>
                      <div className="form-check form-check-inline">
                        <input
                          className="form-check-input"
                          type="radio"
                          name="isBGRelease"
                          value="Yes"
                          id="isBGReleaseYes"
                          checked={isBGRelease === "Yes"}
                          onChange={handleBGReleaseChange}
                          disabled={
                            RelasedBGFiles.length > 0 ||
                            requestClosed === "Yes" ||
                            isUserInAdminM
                          }
                        />
                        <label
                          className="form-check-label"
                          htmlFor="isBGReleaseYes"
                        >
                          Yes
                        </label>
                      </div>
                    </div>
                  </div>

                  {/* Show table ONLY if data exists */}
                  {RelasedBGFiles.length > 0 && (
                    <div className="mt-4">
                      <table className="table table-bordered">
                        <thead>
                          <tr>
                            <th>S.No</th>
                            {/* <th>FileID</th>
                                                    <th>DocID</th> */}
                            <th>Document Name</th>
                            <th>BG Relased Date</th>
                            <th>Actions</th>
                          </tr>
                        </thead>
                        <tbody>
                          {RelasedBGFiles.map((file, index) => (
                            <tr key={file.Id}>
                              <td>{index + 1}</td>
                              {/* <td>{file.FileID}</td>
                                                            <td>{file.DocID}</td> */}

                              <td>{file.FileLeafRef}</td>
                              <td>{formatDateForTable(file.BGReleaseDate)}</td>
                              <td>
                                {(() => {
                                  return (
                                    <button
                                      className="btn btn-sm btn-outline-primary"
                                      onClick={(e) => {
                                        const viewUrl = getViewUrl(file);
                                        handleViewBGFile(
                                          e,
                                          viewUrl,
                                          file.FileLeafRef
                                        );
                                      }}
                                    >
                                      {/* View */}
                                      <FontAwesomeIcon icon={faEye} />
                                    </button>
                                  );
                                })()}

                                <button
                                  className="btn btn-sm btn-outline-success ms-2"
                                  onClick={(e) =>
                                    handleBgDownload(e, file.EncodedAbsUrl)
                                  }
                                >
                                  {/* Download */}
                                  <FontAwesomeIcon icon={faFileArrowDown} />
                                </button>
                                {file.UserEmail === currentUserEmail && (
                                  <button
                                    className="btn btn-sm btn-outline-danger ms-2"
                                    onClick={(e) =>
                                      handleDeleteReleasedBGFile(e, file)
                                    }
                                    disabled={
                                      isUserInAdminM || requestClosed === "Yes"
                                    }
                                  >
                                    {/* Delete */}
                                    <FontAwesomeIcon icon={faTrash} />
                                  </button>
                                )}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}

                  {isBGRelease === "Yes" && RelasedBGFiles.length === 0 && (
                    <div className="mb-3">
                      <label className="form-label">Upload File</label>
                      <div className="input-group">
                        <input
                          type="file"
                          className="form-control"
                          onChange={UploadBgRelaseFile}
                          // disabled={isUserInAdminM  || (props.selectedRow && props.selectedRow.isPaymentReceived === "Yes")}
                          disabled={isUserInAdminM || requestClosed === "Yes"}
                        />
                        <button
                          className="btn btn-success"
                          onClick={handleBGRelaseFileUpload}
                          disabled={isUserInAdminM || requestClosed === "Yes"}
                        >
                          {/* Upload */}
                          <FontAwesomeIcon icon={faUpload} />
                        </button>
                      </div>
                    </div>
                  )}
                </Modal.Body>
                <Modal.Footer>
                  <Button variant="danger" onClick={handleCloseBGModal}>
                    Close
                  </Button>
                </Modal.Footer>
              </Modal>
            </div>
          )}

          <div className="mt-4">
            <div
              className="d-flex align-items-center justify-content-between sectionheader"
              onClick={() => setShowOtherAttachments((prev) => !prev)}
              aria-expanded={showOtherAttachments}
              aria-controls="otherAttachmentsCollapse"
            >
              <h5 className="fw-bold headingColor">Other Attachments</h5>
              <button
                type="button"
                className="btn btn-link"
                onClick={() => setShowOtherAttachments((prev) => !prev)}
                aria-expanded={showOtherAttachments}
                aria-controls="otherAttachmentsCollapse"
                style={{ textDecoration: "none", color: "#ffffff" }}
              >
                {showOtherAttachments ? (
                  <FontAwesomeIcon
                    icon={faAngleUp}
                    onClick={() => setShowOtherAttachments((prev) => !prev)}
                    aria-expanded={showOtherAttachments}
                    aria-controls="otherAttachmentsCollapse"
                  />
                ) : (
                  <FontAwesomeIcon
                    icon={faAngleDown}
                    onClick={() => setShowOtherAttachments((prev) => !prev)}
                    aria-expanded={showOtherAttachments}
                    aria-controls="otherAttachmentsCollapse"
                  />
                )}
              </button>
            </div>
            <div
              className={`${
                showOtherAttachments ? "collapse show" : "collapse"
              } sectioncontent`}
              id="otherAttachmentsCollapse"
            >
              {requestClosed !== "Yes" && !isUserInAdminM ? (
                <div className="row">
                  <div className="col-md-3 mb-3">
                    <label className="form-label">Attachment Type</label>
                    <select
                      disabled={isUserInAdminM || requestClosed == "Yes"}
                      className={`form-select ${
                        errors.attachmentType ? "is-invalid" : ""
                      }`}
                      name="attachmentType"
                      value={formData.attachmentType}
                      onChange={handleTextFieldChange}
                    >
                      <option value="">Select</option>
                      {(isUserInGroupM
                        ? attachmentTypesFinance
                        : attachmentTypes
                      ).map((option) => (
                        <option key={option} value={option}>
                          {option}
                        </option>
                      ))}
                    </select>
                    {errors.attachmentType && (
                      <div className="invalid-feedback">
                        {errors.attachmentType}
                      </div>
                    )}
                  </div>
                  <div className="col-md-4 mb-7">
                    <label className="form-label">Comment</label>
                    <textarea
                      className={`form-control ${
                        errors.comment ? "is-invalid" : ""
                      }`}
                      name="comment"
                      rows={2}
                      value={formData.comment}
                      disabled={isUserInAdminM || requestClosed === "Yes"}
                      onChange={handleTextFieldChange}
                    />
                    {errors.comment && (
                      <div className="invalid-feedback">{errors.comment}</div>
                    )}
                  </div>
                  <div className="col-md-5 mb-3">
                    <label className="form-label" />
                    <div className="input-group">
                      <input
                        type="file"
                        className="form-control"
                        name="file"
                        id="fileInputPO"
                        onChange={handleFileChange}
                        // disabled={isUserInAdminM || RequestClosed == "Yes"}
                        disabled={isUserInAdminM || requestClosed === "Yes"}
                      />
                      <button
                        type="button"
                        className="btn btn-success"
                        onClick={uploadPOAttachment}
                        disabled={isUserInAdminM || requestClosed === "Yes"}
                      >
                        {/* Upload */}
                        <FontAwesomeIcon icon={faUpload} />
                      </button>
                    </div>
                  </div>
                </div>
              ) : null}

              <div className="row mt-4">
                <div className="col-12">
                  {uploadedAttachmentFiles.length > 0 ? (
                    <div className="card">
                      <div className="card-body p-0">
                        <table className="table table-striped table-bordered mb-0">
                          <thead className="table-light">
                            <tr>
                              <th scope="col">S.No</th>
                              <th scope="col">Document Name</th>{" "}
                              {/* FileLeafRef */}
                              <th scope="col">Attachment Type</th>
                              <th scope="col">Comment</th>
                              <th scope="col">Action</th>
                            </tr>
                          </thead>
                          <tbody>
                            {uploadedAttachmentFiles.map((file, index) => (
                              <tr key={file.Id}>
                                <td>{index + 1}</td>
                                <td>{file.FileLeafRef}</td>
                                <td>{file.AttachmentType}</td>
                                <td>{file.Comment}</td>
                                <td>
                                  <button
                                    type="button"
                                    className="btn btn-sm btn-outline-primary me-2"
                                    onClick={(e) => {
                                      const viewUrl = getViewUrl(file);
                                      handlePoAttachmentView(
                                        e,
                                        viewUrl,
                                        file.FileLeafRef
                                      );
                                    }}
                                  >
                                    <FontAwesomeIcon icon={faEye} />{" "}
                                  </button>
                                  <button
                                    type="button"
                                    className="btn btn-sm btn-outline-success me-2"
                                    onClick={(e) =>
                                      handlePoAttachmentDownload(
                                        e,
                                        file.EncodedAbsUrl
                                      )
                                    }
                                  >
                                    <FontAwesomeIcon icon={faFileArrowDown} />{" "}
                                  </button>

                                  {/* {props.selectedRow?.UserEmail === currentUserEmail && ( */}
                                  {file.UserEmail === currentUserEmail &&
                                    requestClosed !== "Yes" && (
                                      <button
                                        type="button"
                                        className="btn btn-sm btn-outline-danger"
                                        onClick={() =>
                                          handlePoAttachmentDelete(file)
                                        }
                                        disabled={
                                          isUserInAdminM ||
                                          requestClosed === "Yes"
                                        }
                                      >
                                        <FontAwesomeIcon icon={faTrash} />
                                      </button>
                                    )}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  ) : (
                    // <div className="alert alert-info mt-2">No Attachment</div>
                    (requestClosed !== "No" || isUserInAdminM) && (
                      <div className="alert alert-info mt-2">No Attachment</div>
                    )
                  )}

                  {/* } */}
                </div>
              </div>
            </div>
          </div>

          {formData.productServiceType?.toLowerCase() === "azure" ? (
            <AzureSection
              siteUrl={siteUrl}
              context={context}
              userGroups={props.userGroups}
              onDataChange={handleAzureSectionDataChange}
              productServiceType={formData.productServiceType}
              customerDetails={customerDetails}
              selectedCustomerName={formData.customerName}
              approverStatus={formData.approverStatus}
              isCollapsed={isAzureSectionCollapsed}
              setIsCollapsed={setIsAzureSectionCollapsed}
              // Only pass these props if rowEdit === "Yes"
              {...(props.rowEdit === "Yes" && {
                rowEdit: props.rowEdit || "",
                selectedRow: props.selectedRow || "",
                setInvoiceRows: props.selectedRow.invoiceDetails || "",
              })}
              currentUserEmail={currentUserEmail}
              onProceedButtonCountChange={setProceedButtonCount}
            />
          ) : (
            <RequesterInvoiceSection
              userGroups={props.userGroups}
              // invoiceRows={invoiceRows}
              invoiceRows={props.rowEdit === "Yes" ? invoiceRows : invoiceRows}
              // invoiceRows={props.rowEdit === "Yes" ? invoiceRows : invoiceRows}
              setInvoiceRows={
                setInvoiceRows as unknown as React.Dispatch<
                  React.SetStateAction<any[]>
                >
              }
              handleInvoiceChange={handleInvoiceChange}
              addInvoiceRow={addInvoiceRow}
              totalPoAmount={Number(formData.poAmount) || 0}
              errors={errors}
              isEditMode={props.rowEdit === "Yes"}
              approverStatus={formData.approverStatus}
              currentUserEmail={currentUserEmail}
              siteUrl={siteUrl}
              context={context}
              onProceedButtonCountChange={setProceedButtonCount}
              hideAddInvoiceButton={["Services"].includes(
                formData.productServiceType
              )}
              isCollapsed={isInvoiceSectionCollapsed}
              setIsCollapsed={setIsInvoiceSectionCollapsed}
              poAmount={Number(formData.poAmount)}
              startDate={formData.startDate}
              endDate={formData.endDate}
              disableDeleteInvoiceRow={
                formData.productServiceType?.toLowerCase() === "resource"
              }
              selectedRowDetails={props.selectedRow}
              props={props} // Pass props to RequesterInvoiceSection
              invoiceApprovalChecked={approvalChecks.invoice}
              setInvoiceApprovalChecked={(v: React.SetStateAction<boolean>) => {
                if (typeof v === "function") {
                  // v is (prev: boolean) => boolean; apply it to current invoice value
                  setApprovalChecks((prev) => ({
                    ...prev,
                    invoice: (v as (prev: boolean) => boolean)(prev.invoice),
                  }));
                } else {
                  // v is a boolean
                  setApprovalChecks((prev) => ({ ...prev, invoice: v }));
                }
              }}
              setDeletedInvoiceItemIDs={setDeletedInvoiceItemIDs}
            />
          )}

          {/* -------------------- Operational Edit Requests Section Start-------------------- */}
          {/* {props.rowEdit === "Yes" && props.selectedRow?.approverStatus ? ( */}
          {props.rowEdit === "Yes" && operationalEdits.length > 0 ? (
            <div className="mt-4">
              <div
                className="d-flex align-items-center justify-content-between sectionheader"
                onClick={() => {
                  setShowOperationalEdits((prev) => !prev);
                  if (!showOperationalEdits) {
                    fetchOperationalEdits(props.selectedRow?.id);
                  }
                }}
                aria-expanded={showOperationalEdits}
                aria-controls="operationalEditsCollapse"
              >
                <h5 className="fw-bold headingColor">Edit Requests History</h5>

                <button
                  type="button"
                  className="btn btn-link"
                  onClick={() => {
                    setShowOperationalEdits((prev) => !prev);
                    if (!showOperationalEdits) {
                      fetchOperationalEdits(props.selectedRow?.id);
                    }
                  }}
                  aria-expanded={showOperationalEdits}
                  aria-controls="operationalEditsCollapse"
                  style={{ textDecoration: "none", color: "#ffffff" }}
                >
                  {showOperationalEdits ? (
                    <FontAwesomeIcon
                      icon={faAngleUp}
                      onClick={() => {
                        setShowOperationalEdits((prev) => !prev);
                        if (!showOperationalEdits) {
                          fetchOperationalEdits(props.selectedRow?.id);
                        }
                      }}
                      aria-expanded={showOperationalEdits}
                      aria-controls="operationalEditsCollapse"
                    />
                  ) : (
                    <FontAwesomeIcon
                      icon={faAngleDown}
                      onClick={() => {
                        setShowOperationalEdits((prev) => !prev);
                        if (!showOperationalEdits) {
                          fetchOperationalEdits(props.selectedRow?.id);
                        }
                      }}
                      aria-expanded={showOperationalEdits}
                      aria-controls="operationalEditsCollapse"
                    />
                  )}
                </button>
              </div>

              <div
                className={`${
                  showOperationalEdits ? "collapse show" : "collapse"
                } sectioncontent`}
                id="operationalEditsCollapse"
              >
                {loadingOperationalEdits ? (
                  <div className="text-center my-3">
                    <Spinner animation="border" size="sm" /> Loading...
                  </div>
                ) : (
                  <>
                    {/* Current / Old toggles */}
                    <div className="mb-3 d-flex gap-2">
                      <button
                        type="button"
                        className={`btn btn-sm ${
                          showCurrentOperational
                            ? "btn-primary"
                            : "btn-outline-secondary"
                        }`}
                        onClick={() => {
                          setShowCurrentOperational(true);
                          setShowOldOperational(false);
                          fetchOperationalEdits(props.selectedRow?.id);
                        }}
                      >
                        Current Requests
                      </button>
                      <button
                        type="button"
                        className={`btn btn-sm ${
                          showOldOperational
                            ? "btn-primary"
                            : "btn-outline-secondary"
                        }`}
                        onClick={() => {
                          setShowOldOperational(true);
                          setShowCurrentOperational(false);
                          fetchOperationalEdits(props.selectedRow?.id);
                        }}
                      >
                        All Requests
                      </button>
                    </div>

                    {/* prepare lists */}
                    {(() => {
                      const currentRequests = operationalEdits.filter(
                        (r: any) => {
                          const st = (r.Status || "").trim();
                          return (
                            st === "Pending Approval" ||
                            st === "Hold" ||
                            st === "Reminder"
                          );
                        }
                      );
                      const oldRequests = operationalEdits.filter((r: any) => {
                        const os = (r.Status || "").trim();
                        return (
                          os !== "Pending Approval" ||
                          os !== "Hold" ||
                          os !== "Reminder"
                        );
                      });

                      const isProjectManager =
                        (currentUserEmail || "").toLowerCase() ===
                        (
                          props.selectedRow?.projectMangerEmail || ""
                        ).toLowerCase();
                      const isEmployee =
                        (currentUserEmail || "").toLowerCase() ===
                        (props.selectedRow?.employeeEmail || "").toLowerCase();

                      const renderRows = (
                        rows: any[],
                        showActions: boolean
                      ) => (
                        <div className="table-responsive card">
                          <div style={{ maxHeight: 300, overflowY: "auto" }}>
                            <table className="table table-hover mb-0">
                              <thead
                                className="table-light"
                                style={{
                                  position: "sticky",
                                  top: 0,
                                  zIndex: 2,
                                }}
                              >
                                <tr>
                                  <th style={{ width: 40 }}>#</th>
                                  <th>Reason</th>
                                  <th>Selected Sections</th>
                                  <th style={{ width: 180 }}>User</th>
                                  <th style={{ width: 150 }}>Created Date</th>
                                  <th style={{ width: 120 }}>Status</th>
                                  {showActions && (
                                    <th
                                      className="text-center"
                                      style={{ width: 220 }}
                                    >
                                      Actions
                                    </th>
                                  )}
                                </tr>
                              </thead>
                              <tbody>
                                {rows.map((row: any, idx: number) => (
                                  <tr key={row.Id || idx}>
                                    <td style={{ verticalAlign: "middle" }}>
                                      {idx + 1}
                                    </td>
                                    <td
                                      style={{
                                        maxWidth: 240,
                                        wordBreak: "break-word",
                                      }}
                                    >
                                      {row.Reason || "-"}
                                    </td>
                                    <td>
                                      {(row.SelectedSections || "")
                                        .split(",")
                                        .map((s: string, i: number) => {
                                          const txt = s.trim();
                                          const badgeClass = txt
                                            .toLowerCase()
                                            .includes("client")
                                            ? "bg-primary"
                                            : txt.toLowerCase().includes("po")
                                            ? "bg-warning text-dark"
                                            : txt
                                                .toLowerCase()
                                                .includes("invoice")
                                            ? "bg-success"
                                            : "bg-secondary";
                                          return (
                                            <span
                                              key={i}
                                              className={`badge ${badgeClass} me-1`}
                                              style={{ fontSize: 12 }}
                                            >
                                              {txt || "-"}
                                            </span>
                                          );
                                        })}
                                    </td>
                                    <td>{row.UserName || "-"}</td>
                                    <td>
                                      {row.Created
                                        ? new Date(
                                            row.Created
                                          ).toLocaleDateString("en-GB")
                                        : "-"}
                                    </td>
                                    <td>
                                      <span
                                        className={`badge ${
                                          row.Status === "Pending Approval"
                                            ? "bg-info"
                                            : row.Status === "Approved"
                                            ? "bg-success"
                                            : row.Status === "Reject"
                                            ? "bg-danger"
                                            : "bg-secondary"
                                        }`}
                                      >
                                        {row.Status || "N/A"}
                                      </span>
                                    </td>

                                    {showActions && (
                                      <td className="text-center">
                                        <div
                                          className="btn-group btn-group-sm"
                                          role="group"
                                        >
                                          {/* Approve/Hold/Reject only visible to Project Manager */}
                                          {requestClosed !== "Yes" &&
                                            isProjectManager &&
                                            (row.Status ===
                                              "Pending Approval" ||
                                              row.Status === "Hold" ||
                                              row.Status === "Reminder") && (
                                              <>
                                                <button
                                                  className="btn btn-success"
                                                  onClick={async (
                                                    e: React.MouseEvent<HTMLButtonElement>
                                                  ) => {
                                                    e.preventDefault();
                                                    if (!row.Id) return;
                                                    try {
                                                      setIsLoading(true);

                                                      // Initialize EditRequestData with default values
                                                      let EditRequestData: Record<
                                                        string,
                                                        any
                                                      > = {
                                                        RunWF: "Yes",
                                                        Status: "Approved",
                                                        ReminderDate: todayDate,
                                                      };
                                                      let shouldUpdateMainList =
                                                        false;

                                                      console.log(
                                                        "Approving Row----",
                                                        row
                                                      );

                                                      // Handle PO Section
                                                      if (
                                                        row?.SelectedSections?.includes(
                                                          "po"
                                                        )
                                                      ) {
                                                        EditRequestData = {
                                                          ...EditRequestData,
                                                          PoNo: formData.poNo,
                                                          PoDate:
                                                            formData.poDate
                                                              ? moment(
                                                                  formData.poDate,
                                                                  "DD-MM-YYYY"
                                                                ).format(
                                                                  "YYYY-MM-DD"
                                                                )
                                                              : null,
                                                          POAmount: Number(
                                                            formData.poAmount
                                                          ),
                                                          BgDate:
                                                            formData.bgDate
                                                              ? moment(
                                                                  formData.bgDate,
                                                                  "DD-MM-YYYY"
                                                                ).format(
                                                                  "YYYY-MM-DD"
                                                                )
                                                              : null,
                                                        };
                                                      }

                                                      // Handle Client Section
                                                      if (
                                                        row?.SelectedSections?.includes(
                                                          "client"
                                                        )
                                                      ) {
                                                        EditRequestData = {
                                                          ...EditRequestData,
                                                          CustomerEmail:
                                                            formData.customerEmail,
                                                          Location:
                                                            formData.location,
                                                          WorkTitle:
                                                            formData.workTitle,
                                                          WorkDetails:
                                                            formData.workDetail,
                                                        };
                                                      }

                                                      // Handle Invoice Section
                                                      if (
                                                        row?.SelectedSections?.includes(
                                                          "invoice"
                                                        )
                                                      ) {
                                                        for (const invoiceRow of invoiceRows) {
                                                          if (
                                                            invoiceRow?.InvoiceStatus ===
                                                            "Pending Approval"
                                                          ) {
                                                            const invoiceDetails =
                                                              {
                                                                Comments:
                                                                  invoiceRow.InvoiceDescription,
                                                                InvoiceAmount:
                                                                  Number(
                                                                    invoiceRow.InvoiceAmount
                                                                  ),
                                                                InvoiceDueDate:
                                                                  invoiceRow.InvoiceDueDate
                                                                    ? moment(
                                                                        invoiceRow.InvoiceDueDate,
                                                                        "DD-MM-YYYY"
                                                                      ).format(
                                                                        "YYYY-MM-DD"
                                                                      )
                                                                    : null,
                                                                RequestID:
                                                                  invoiceRow.RequestID,
                                                                ClaimNo:
                                                                  invoiceRow?.ClaimNo ||
                                                                  "",
                                                                DocId:
                                                                  invoiceRow?.DocId ||
                                                                  "",

                                                                InvoiceDate:
                                                                  invoiceRow.InvoiceDate
                                                                    ? moment(
                                                                        invoiceRow.InvoiceDate,
                                                                        "DD-MM-YYYY"
                                                                      ).format(
                                                                        "YYYY-MM-DD"
                                                                      )
                                                                    : null,
                                                                InvoiceFileID:
                                                                  invoiceRow?.InvoiceFileID ||
                                                                  "",
                                                                InvoicNo:
                                                                  invoiceRow?.InvoiceNo ||
                                                                  "",
                                                                InvoiceStatus:
                                                                  invoiceRow?.InvoiceStatus ||
                                                                  "",
                                                                InvoiceTaxAmount:
                                                                  Number(
                                                                    invoiceRow?.InvoiceTaxAmount
                                                                  ) || 0,
                                                                PendingAmount:
                                                                  Number(
                                                                    invoiceRow?.PendingAmount
                                                                  ) || 0,
                                                                PrevInvoiceStatus:
                                                                  invoiceRow?.PrevInvoiceStatus ||
                                                                  "",
                                                                PoAmount:
                                                                  Number(
                                                                    invoiceRow?.RemainingPoAmount
                                                                  ) || 0,
                                                                ContractID:
                                                                  formData.requestId ||
                                                                  "",
                                                                EditRequestItemID:
                                                                  row.Id || "",
                                                              };

                                                            // Save invoice details to SharePoint
                                                            await saveDataToSharePoint(
                                                              OperationalEditInvoiceHistory,
                                                              invoiceDetails,
                                                              siteUrl
                                                            );

                                                            if (
                                                              invoiceRow?.PrevInvoiceStatus ===
                                                              "Generated"
                                                            ) {
                                                              // ✅ Also update this in the main InvoiceList
                                                              if (
                                                                invoiceRow?.itemID
                                                              ) {
                                                                try {
                                                                  await updateDataToSharePoint(
                                                                    InvoicelistName, // Your main invoice list name
                                                                    {
                                                                      CreditNoteStatus:
                                                                        "Pending",
                                                                    },
                                                                    siteUrl,
                                                                    invoiceRow.itemID
                                                                  );
                                                                  console.log(
                                                                    `CreditNoteStatus updated to "Pending" for invoice ID ${invoiceRow.itemID}`
                                                                  );

                                                                  shouldUpdateMainList =
                                                                    true;
                                                                } catch (updateError) {
                                                                  console.error(
                                                                    "Error updating CreditNoteStatus in main InvoiceList:",
                                                                    updateError
                                                                  );
                                                                }
                                                              }
                                                            }
                                                          }
                                                        }
                                                      }

                                                      // Update Operational Edit Request status
                                                      await updateDataToSharePoint(
                                                        OperationalEditRequest,
                                                        EditRequestData,
                                                        siteUrl,
                                                        row.Id
                                                      );

                                                      // Mirror status to Main list
                                                      if (
                                                        props.selectedRow?.id
                                                      ) {
                                                        await updateDataToSharePoint(
                                                          MainList,
                                                          {
                                                            ApproverStatus:
                                                              "Approved",
                                                            RunWF: "Yes",
                                                            ...(shouldUpdateMainList && {
                                                              IsCreditNoteUploaded:
                                                                "No",
                                                            }),
                                                          },
                                                          siteUrl,
                                                          props.selectedRow.id
                                                        );
                                                      }

                                                      // Show success message and refresh data
                                                      showSnackbar(
                                                        "Request approved.",
                                                        "success"
                                                      );
                                                      await fetchOperationalEdits(
                                                        props.selectedRow?.id
                                                      );
                                                      await props.refreshCmsDetails?.();
                                                      await finalizeAction(
                                                        false
                                                      );
                                                    } catch (err) {
                                                      console.error(
                                                        "Error approving request:",
                                                        err
                                                      );
                                                      showSnackbar(
                                                        "Failed to approve request.",
                                                        "error"
                                                      );
                                                    } finally {
                                                      setIsLoading(false);
                                                    }
                                                  }}
                                                >
                                                  Approve
                                                </button>

                                                <button
                                                  className="btn btn-warning"
                                                  onClick={async (
                                                    e: React.MouseEvent<HTMLButtonElement>
                                                  ) => {
                                                    e.preventDefault();
                                                    if (!row.Id) return;
                                                    try {
                                                      setIsLoading(true);
                                                      // set OperationalEditRequest -> Hold
                                                      await updateDataToSharePoint(
                                                        OperationalEditRequest,
                                                        {
                                                          Status: "Hold",
                                                          ReminderDate:
                                                            todayDate,
                                                        },
                                                        siteUrl,
                                                        row.Id
                                                      );
                                                      // mirror Hold to main list ApproverStatus
                                                      if (
                                                        props.selectedRow?.id
                                                      ) {
                                                        await updateDataToSharePoint(
                                                          MainList,
                                                          {
                                                            ApproverStatus:
                                                              "Hold",
                                                            RunWF: "Yes",
                                                          },
                                                          siteUrl,
                                                          props.selectedRow.id
                                                        );
                                                      }
                                                      showSnackbar(
                                                        "Request placed on hold.",
                                                        "info"
                                                      );
                                                      await fetchOperationalEdits(
                                                        props.selectedRow?.id
                                                      );
                                                      await props.refreshCmsDetails?.();
                                                      await finalizeAction(
                                                        false
                                                      );
                                                    } catch (err) {
                                                      console.error(err);
                                                      showSnackbar(
                                                        "Failed to put request on hold.",
                                                        "error"
                                                      );
                                                    } finally {
                                                      setIsLoading(false);
                                                    }
                                                  }}
                                                >
                                                  Hold
                                                </button>

                                                {/* <button
                                                  className="btn btn-danger"
                                                  onClick={async () => {
                                                    if (!row.Id) return;
                                                    if (
                                                      !window.confirm(
                                                        "Reject this edit request?"
                                                      )
                                                    )
                                                      return;
                                                    try {
                                                      setIsLoading(true);
                                                      // mark OperationalEditRequest rejected
                                                      await updateDataToSharePoint(
                                                        OperationalEditRequest,
                                                        { Status: "Rejected" },
                                                        siteUrl,
                                                        row.Id
                                                      );
                                                      // mirror rejection to main list (use "Reject" to match existing main list values)
                                                      if (
                                                        props.selectedRow?.id
                                                      ) {
                                                        await updateDataToSharePoint(
                                                          MainList,
                                                          {
                                                            ApproverStatus:
                                                              "Reject",
                                                            RunWF: "No",
                                                          },
                                                          siteUrl,
                                                          props.selectedRow.id
                                                        );
                                                      }
                                                      showSnackbar(
                                                        "Request rejected.",
                                                        "success"
                                                      );
                                                      await fetchOperationalEdits(
                                                        props.selectedRow?.id
                                                      );
                                                      await props.refreshCmsDetails?.();
                                                      await finalizeAction(
                                                        false
                                                      );
                                                    } catch (err) {
                                                      console.error(err);
                                                      showSnackbar(
                                                        "Failed to reject request.",
                                                        "error"
                                                      );
                                                    } finally {
                                                      setIsLoading(false);
                                                    }
                                                  }}
                                                >
                                                  Reject
                                                </button> */}

                                                <button
                                                  className="btn btn-danger"
                                                  onClick={async (
                                                    e: React.MouseEvent<HTMLButtonElement>
                                                  ) => {
                                                    e.preventDefault();
                                                    if (!row.Id) return;
                                                    if (
                                                      !window.confirm(
                                                        "Reject this edit request?"
                                                      )
                                                    )
                                                      return;
                                                    try {
                                                      setIsLoading(true);

                                                      // Mark OperationalEditRequest as rejected
                                                      await updateDataToSharePoint(
                                                        OperationalEditRequest,
                                                        {
                                                          Status: "Reject",
                                                          RunWF: "Yes",
                                                          ReminderDate:
                                                            todayDate,
                                                        },
                                                        siteUrl,
                                                        row.Id
                                                      );

                                                      // Update InvoiceStatus and PrevInvoiceStatus for "Pending Approval" invoices
                                                      for (const invoiceRow of invoiceRows) {
                                                        if (
                                                          invoiceRow?.InvoiceStatus ===
                                                          "Pending Approval"
                                                        ) {
                                                          const invoiceUpdateData =
                                                            {
                                                              InvoiceStatus:
                                                                invoiceRow.PrevInvoiceStatus ||
                                                                "",
                                                              PrevInvoiceStatus:
                                                                "",
                                                            };

                                                          await updateDataToSharePoint(
                                                            InvoicelistName,
                                                            invoiceUpdateData,
                                                            siteUrl,
                                                            Number(
                                                              invoiceRow.itemID
                                                            )
                                                          );
                                                        }
                                                      }

                                                      // Mirror rejection to the main list
                                                      if (
                                                        props.selectedRow?.id
                                                      ) {
                                                        const updatedData = {
                                                          ApproverStatus:
                                                            "Reject",
                                                          RunWF: "Yes",
                                                        };
                                                        await updateDataToSharePoint(
                                                          MainList,
                                                          updatedData,
                                                          siteUrl,
                                                          props.selectedRow.id
                                                        );
                                                      }

                                                      showSnackbar(
                                                        "Request rejected.",
                                                        "success"
                                                      );
                                                      await fetchOperationalEdits(
                                                        props.selectedRow?.id
                                                      );
                                                      await props.refreshCmsDetails?.();
                                                      await finalizeAction(
                                                        false
                                                      );
                                                    } catch (err) {
                                                      console.error(err);
                                                      showSnackbar(
                                                        "Failed to reject request.",
                                                        "error"
                                                      );
                                                    } finally {
                                                      setIsLoading(false);
                                                    }
                                                  }}
                                                >
                                                  Reject
                                                </button>
                                              </>
                                            )}

                                          {/* Remind button: employee only, but not when same as project manager */}
                                          {requestClosed !== "Yes" &&
                                            isEmployee &&
                                            !isProjectManager &&
                                            (row.Status ===
                                              "Pending Approval" ||
                                              row.Status === "Hold" ||
                                              row.Status === "Reminder") && (
                                              <button
                                                className="btn btn-warning"
                                                onClick={async (
                                                  e: React.MouseEvent<HTMLButtonElement>
                                                ) => {
                                                  e.preventDefault();
                                                  if (!props.selectedRow?.id)
                                                    return;
                                                  handleReminder(
                                                    e as any,
                                                    props.selectedRow.id
                                                  );
                                                }}
                                              >
                                                <FontAwesomeIcon
                                                  icon={faBell}
                                                />{" "}
                                                Remind
                                              </button>
                                            )}
                                        </div>
                                      </td>
                                    )}
                                  </tr>
                                ))}
                                {rows.length === 0 && (
                                  <tr>
                                    <td
                                      colSpan={showActions ? 6 : 5}
                                      className="text-center py-3"
                                    >
                                      No requests
                                    </td>
                                  </tr>
                                )}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      );

                      return (
                        <>
                          {showCurrentOperational && (
                            <div className="mb-3">
                              <h6 className="mb-2">
                                Current Requests (Pending Approval)
                              </h6>
                              {renderRows(currentRequests, true)}
                            </div>
                          )}
                          {showOldOperational && (
                            <div className="mb-3">
                              <h6 className="mb-2">All Requests</h6>
                              {renderRows(oldRequests, false)}
                            </div>
                          )}
                        </>
                      );
                    })()}
                  </>
                )}
              </div>
            </div>
          ) : null}
          {/* -------------------- Operational Edit Requests End -------------------- */}

          <div className="row d-flex justify-content-end">
            <div className="col-md-6 d-flex justify-content-end">
              {/* Show Update button if rowEdit is "Yes" and employeeEmail matches currentUserEmail */}
              <button
                type="button"
                className="btn btn-danger w-40 mt-3 me-2"
                onClick={handleExit}
              >
                <FontAwesomeIcon icon={faXmark} /> Exit
              </button>
              {props.rowEdit === "Yes" &&
              props.selectedRow?.employeeEmail === currentUserEmail ? (
                <>
                  {(formData.approverStatus === "Completed" ||
                    formData.approverStatus === "" ||
                    formData.approverStatus === "Reminder" ||
                    formData.approverStatus === "Reject") &&
                    requestClosed !== "Yes" &&
                    proceedButtonCount > 0 && (
                      <>
                        <textarea
                          className={`form-control ${
                            errors.editReason ? "is-invalid" : ""
                          }`}
                          name="editReason"
                          rows={2}
                          placeholder="Enter reason for edit request"
                          // value={formData.editReason}
                          disabled={isUserInAdminM || requestClosed === "Yes"}
                          onChange={handleTextFieldChange}
                          style={{
                            // backgroundColor: "#107c10",

                            display: "none",
                          }}
                        />
                      </>
                    )}

                  {props.rowEdit === "Yes" &&
                    props.selectedRow?.employeeEmail === currentUserEmail &&
                    requestClosed !== "Yes" &&
                    ![
                      "Approved",
                      "Hold",
                      "Pending From Approver",
                      "Reminder",
                    ].includes(props.selectedRow.approverStatus) &&
                    props.selectedRow?.isCreditNoteUploaded !== "No" && (
                      <button
                        type="button"
                        className="btn btn-primary w-40 mt-3"
                        onClick={(e) => {
                          e.preventDefault();
                          if (
                            !approvalChecks.client &&
                            !approvalChecks.po &&
                            !approvalChecks.invoice
                          ) {
                            showSnackbar(
                              "Please select at least one section (Client & Work / PO / Invoice) to approve the edit.",
                              "error"
                            );
                            return;
                          }
                          // reuse existing edit handler
                          handleEditRequestApproval(e, props.selectedRow.id);
                        }}
                        style={{ marginLeft: "10px", marginRight: "10px" }}
                      >
                        <FontAwesomeIcon icon={faEdit} /> Edit Request Approval
                      </button>
                    )}

                  {/* {props.rowEdit === "Yes" &&
                    invoiceRows.some(
                      (row) => row.invoiceCloseApprovalChecked
                    ) && (
                      <button
                        type="button"
                        className="btn btn-primary w-40 mt-3"
                        onClick={(e) => {
                          e.preventDefault();
                          // alert("Close PO functionality triggered!");
                          handleInvoiceClose(e, invoiceRows);
                        }}
                        style={{ marginLeft: "10px", marginRight: "10px" }}
                      >
                        <FontAwesomeIcon icon={faEdit} /> Close Po
                      </button>
                    )} */}

                  {props.rowEdit === "Yes" &&
                    invoiceRows.some(
                      (row) => row.invoiceCloseApprovalChecked
                    ) && (
                      <button
                        type="button"
                        className="btn btn-primary w-40 mt-3"
                        onClick={(e) => {
                          e.preventDefault();
                          handleInvoiceClose(e);
                        }}
                        style={{ marginLeft: "10px", marginRight: "10px" }}
                      >
                        <FontAwesomeIcon icon={faEdit} /> Close Po
                      </button>
                    )}

                  {props.rowEdit === "Yes" &&
                    props.selectedRow?.employeeEmail === currentUserEmail &&
                    requestClosed !== "Yes" &&
                    props.selectedRow.approverStatus === "Approved" && (
                      <button
                        type="button"
                        className="btn btn-primary w-40 mt-3"
                        onClick={(e) => handleUpdateEditRequest(e)}
                        style={{ marginLeft: "10px", marginRight: "10px" }}
                      >
                        <FontAwesomeIcon icon={faEdit} /> Update Edit Request
                        Approval
                      </button>
                    )}
                  <Modal
                    show={isPopupOpen}
                    onHide={handleClosePopup}
                    centered
                    dialogClassName="custom-modal-width"
                  >
                    <Modal.Header closeButton className="custom-modal-header">
                      <Modal.Title>Edit Request Approval</Modal.Title>
                    </Modal.Header>
                    <Modal.Body>
                      {isSubmitting && ( // Loader inside the modal
                        <div
                          style={{
                            position: "absolute",
                            top: 0,
                            left: 0,
                            width: "100%",
                            height: "100%",
                            background: "rgba(255, 255, 255, 0.8)",
                            zIndex: 1050,
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "center",
                          }}
                        >
                          <Spinner animation="border" variant="primary" />
                          <span className="ms-3">Submitting...</span>
                        </div>
                      )}
                      <div
                        style={{
                          display: "flex",
                          flexDirection: "column",
                          gap: "10px",
                        }}
                      >
                        <div>
                          <label>
                            <strong>Selected Section:</strong>
                          </label>
                          <p>
                            {[
                              approvalChecks.client && "Client & Work Detail",
                              approvalChecks.po && "PO Details",
                              approvalChecks.invoice && "Invoice Details",
                            ]
                              .filter(Boolean)
                              .join(", ")}
                          </p>
                        </div>
                        <div>
                          <label htmlFor="reason">
                            <strong>
                              Reason:<span style={{ color: "red" }}>*</span>
                            </strong>
                          </label>
                          <textarea
                            id="reason"
                            className="form-control"
                            value={reason}
                            onChange={(e) => setReason(e.target.value)}
                            placeholder="Enter your reason here..."
                            rows={4}
                          />
                        </div>
                      </div>
                    </Modal.Body>
                    <Modal.Footer className="d-flex justify-content-between">
                      <Button
                        variant="secondary"
                        onClick={handleClosePopup}
                        disabled={isSubmitting}
                      >
                        Close
                      </Button>

                      <Button
                        variant="success"
                        disabled={isSubmitting}
                        onClick={(e) => {
                          if (selectedId !== null) {
                            handleSubmitEditRequestApproval(e, selectedId);
                          } else {
                            setModalSnackbar({
                              open: true,
                              message: "No request selected for approval.",
                              severity: "error",
                            });
                          }
                        }}
                      >
                        Submit
                      </Button>
                    </Modal.Footer>
                    <Snackbar
                      open={modalSnackbar.open}
                      autoHideDuration={6000}
                      onClose={() =>
                        setModalSnackbar({ ...modalSnackbar, open: false })
                      }
                      anchorOrigin={{ vertical: "top", horizontal: "center" }}
                    >
                      <Alert
                        onClose={() =>
                          setModalSnackbar({ ...modalSnackbar, open: false })
                        }
                        severity={
                          ["success", "error", "warning", "info"].includes(
                            modalSnackbar.severity
                          )
                            ? (modalSnackbar.severity as
                                | "success"
                                | "error"
                                | "warning"
                                | "info")
                            : "info" // Default to 'info' if the value is invalid
                        }
                      >
                        {modalSnackbar.message}
                      </Alert>
                    </Snackbar>
                  </Modal>
                  {/* <Modal
                    show={isPopupOpen}
                    onHide={handleClosePopup}
                    centered
                    dialogClassName="custom-modal-width"
                  >
                    <Modal.Header closeButton className="custom-modal-header">
                      <Modal.Title>Edit Request Approval</Modal.Title>
                    </Modal.Header>
                    <Modal.Body>
                      <div
                        style={{
                          display: "flex",
                          flexDirection: "column",
                          gap: "10px",
                        }}
                      >
                        <div>
                          <label>
                            <strong>Selected Section:</strong>
                          </label>
                          <p>
                         

                            {[
                              approvalChecks.client && "Client & Work Detail",
                              approvalChecks.po && "PO Details",
                              approvalChecks.invoice && "Invoice Details",
                            ]
                              .filter(Boolean) // Remove any `false` or `undefined` values
                              .join(", ")}
                          </p>
                        </div>

                        <div style={{ display: "none" }}>
                          <label>
                            <strong>Selected ID:</strong>
                          </label>
                          <p>{selectedId}</p> 
                        </div>

                        <div>
                          <label htmlFor="reason">
                            <strong>
                              Reason:<span style={{ color: "red" }}>*</span>
                            </strong>
                          </label>
                          <textarea
                            id="reason"
                            className="form-control"
                            value={reason}
                            onChange={(e) => setReason(e.target.value)}
                            placeholder="Enter your reason here..."
                            rows={4}
                          />
                        </div>
                      </div>
                    </Modal.Body>
                    <Modal.Footer>
                      <Button
                        variant="success"
                        // onClick={() => {
                        //   if (selectedId !== null) {
                        //     handleSubmitEditRequestApproval(selectedId);
                        //   } else {
                        //     alert("No request selected for approval.");
                        //   }
                        // }}
                        onClick={(e) => {
                          if (selectedId !== null) {
                            handleSubmitEditRequestApproval(e, selectedId);
                          } else {
                            alert("No request selected for approval.");
                          }
                        }}
                      >
                        <FontAwesomeIcon icon={faPaperPlane} /> Submit
                      </Button>
                      <Button variant="danger" onClick={handleClosePopup}>
                        Close
                      </Button>
                    </Modal.Footer>
                  </Modal> */}
                  {/* {requestClosed !== "Yes" ? (
                    <>
                      <button
                        type="button"
                        className="btn btn-info w-40 mt-3"
                        //  onClick={handleProceedApprove()}
                        onClick={(e) =>
                          handleProceedApprove(e, props.selectedRow.id)
                        }
                        // style={"display":"none"}
                        style={{
                          // backgroundColor: "#107c10",
                          marginLeft: "10px",
                          marginRight: "10px",
                        }}
                      >
                        Approve
                      </button>
                      <button
                        type="button"
                        className="btn btn-info w-40 mt-3"
                        // onClick={handleProceedReject()}
                        onClick={(e) =>
                          handleProceedReject(e, props.selectedRow.id)
                        }
                        // style={"display":"none"}
                        style={{
                          // backgroundColor: "#107c10",
                          marginLeft: "10px",
                          marginRight: "10px",
                        }}
                      >
                        Reject
                      </button>
                    </>
                  ) : null} */}

                  {/* {(formData.approverStatus === "Completed" ||
                    formData.approverStatus === "" ||
                    // formData.approverStatus === "Reminder" ||
                    formData.approverStatus === "Reject") &&
                    requestClosed !== "Yes" &&
                    proceedButtonCount > 0 && (
                      <button
                        type="button"
                        className="btn btn-success w-40 mt-3"
                        onClick={(e) =>
                          handleEditRequest(e, props.selectedRow.id)
                        }
                        style={{
                          // backgroundColor: "#107c10",
                          marginLeft: "10px",
                          marginRight: "10px",
                          display: "none",
                        }}
                      >
                        <FontAwesomeIcon icon={faEdit} /> Edit Request
                      </button>
                    )} */}

                  {formData.approverStatus === "Approvd" &&
                    requestClosed !== "Yes" &&
                    proceedButtonCount > 0 && (
                      <a
                        href={`${siteUrl}/SitePages/InvoiceForm.aspx??editable=true&FileID=${props.selectedRow.docID}&itemId=${props.selectedRow.id}`}
                        className="btn btn-warning w-40 mt-3"
                        style={{
                          textDecoration: "none",
                          // color: "#107c10",
                          marginLeft: "10px",
                          marginRight: "10px",
                        }}
                      >
                        <FontAwesomeIcon icon={faEdit} /> Edit the Invoice
                      </a>
                    )}
                  {/* {(formData.approverStatus === "Hold" ||
                    formData.approverStatus === "Reminder") &&
                    requestClosed !== "Yes" &&
                    proceedButtonCount > 0 && (
                      <button
                        type="button"
                        className="btn btn-success w-40 mt-3 me-2"
                        onClick={(e) => handleReminder(e, props.selectedRow.id)}
                      >
                        Reminder
                      </button>
                    )} */}
                </>
              ) : (
                <>
                  {!props.selectedRow ||
                  props.selectedRow.employeeEmail === currentUserEmail ? (
                    <>
                      <button
                        type="submit"
                        className="btn btn-success w-40 mt-3 "
                        disabled={isUserInAdminM || requestClosed === "Yes"}
                        style={{
                          marginLeft: "10px",
                          marginRight: "10px",
                        }}
                      >
                        <FontAwesomeIcon icon={faPaperPlane} /> Submit
                      </button>
                    </>
                  ) : null}
                </>
              )}
              {/* {props.selectedRow &&
                formData.productServiceType.toLowerCase() == "azure" &&
                props.selectedRow.isAzureRequestClosed !== "Yes" && (
                  <>
                    <button
                      type="submit"
                      className="btn btn-success w-40 mt-3 "
                      disabled={isUserInAdminM || requestClosed === "Yes"}
                      style={{
                        // backgroundColor: "#107c10",
                        marginLeft: "10px",
                        marginRight: "10px",
                      }}
                    >
                      <FontAwesomeIcon icon={faPaperPlane} /> Submit
                    </button>
                  </>
                )} */}

              {props.selectedRow &&
              formData.productServiceType.toLowerCase() === "azure" &&
              props.selectedRow.isAzureRequestClosed !== "Yes" &&
              props.selectedRow.employeeEmail === currentUserEmail
                ? (() => {
                    const lastRow =
                      azureSectionData && azureSectionData.length > 0
                        ? azureSectionData[azureSectionData.length - 1]
                        : null;
                    // Only show the submit button when last row exists and does NOT have an ItemID
                    if (lastRow && !lastRow.itemID) {
                      return (
                        <>
                          <button
                            type="submit"
                            className="btn btn-warning w-40 mt-3 "
                            disabled={isUserInAdminM || requestClosed === "Yes"}
                            style={{
                              marginLeft: "10px",
                              marginRight: "10px",
                            }}
                            onClick={handleSubmitAzureDetails}
                          >
                            <FontAwesomeIcon icon={faPaperPlane} /> Submit Azure
                            Details
                          </button>
                        </>
                      );
                    }
                    return null;
                  })()
                : null}
            </div>
          </div>
        </form>
      </div>
    </div>
  );
};

export default RequestForm;
