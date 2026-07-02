<script lang="ts">
  import * as XLSX from "xlsx";
  import { allocateRooms as buildAllocations } from "$lib/allocation";
  import type { Allocation, Participant, Room } from "$lib/types";
  import Building2 from "@lucide/svelte/icons/building-2";
  import Moon from "@lucide/svelte/icons/moon";
  import Sun from "@lucide/svelte/icons/sun";
  import Bed from "@lucide/svelte/icons/bed";
  import Users from "@lucide/svelte/icons/users-round";
  import Checkcircle from "@lucide/svelte/icons/circle-check-big";
  import BedDouble from "@lucide/svelte/icons/bed-double";
  import Cloudupload from "@lucide/svelte/icons/cloud-upload";
  import FileSpreadsheet from "@lucide/svelte/icons/file-spreadsheet";
  import FileUser from "@lucide/svelte/icons/file-user";
  import List from "@lucide/svelte/icons/list";
  import Clock from "@lucide/svelte/icons/clock-3";
  import Mars from "@lucide/svelte/icons/mars";
  import Venus from "@lucide/svelte/icons/venus";
  import Download from "@lucide/svelte/icons/download";

  interface HostelRow {
    "Hostel Name": string;
    "Room Number": string | number;
    Capacity: string | number;
    "Hostel Type": string;
  }

  interface ParticipantRow {
    "Participant Name": string;
    Gender: string;
  }

  interface UploadedFileInfo {
    name: string;
    size: number;
  }

  type PreviewTab = "rooms" | "participants" | "allocation";
  type Theme = "dark" | "light";

  let hostelFileInput = $state<HTMLInputElement | null>(null);
  let participantFileInput = $state<HTMLInputElement | null>(null);
  let hostelFileInfo = $state<UploadedFileInfo | null>(null);
  let participantFileInfo = $state<UploadedFileInfo | null>(null);
  let previewTab = $state<PreviewTab>("rooms");
  let theme = $state<Theme>("dark");

  let rooms = $state<Room[]>([]);
  let participants = $state<Participant[]>([]);
  let allocations: Allocation[] = $state([]);

  let total_capacity = $state(0);
  let remaining_capacity = $state(0);
  let allocated = $state(false);
  let allocated_participants = $state(0);
  let participant_no = $state(0);
  let unallocated_participants = $state(0);
  let lastAllocationAt = $state<string | null>(null);

  const allocationDisabled = $derived(
    rooms.length === 0 || participants.length === 0 || allocated,
  );
  const hasUploadedFiles = $derived(
    rooms.length > 0 || participants.length > 0,
  );
  const hasBothFiles = $derived(rooms.length > 0 && participants.length > 0);
  const canDownload = $derived(allocated && allocations.length > 0);
  const totalHostels = $derived(
    new Set(rooms.map((room) => room.hostelName)).size,
  );
  const boysHostels = $derived(
    new Set(
      rooms
        .filter((room) => normalize(room.hostelType) === "boys")
        .map((room) => room.hostelName),
    ).size,
  );
  const girlsHostels = $derived(
    new Set(
      rooms
        .filter((room) => normalize(room.hostelType) === "girls")
        .map((room) => room.hostelName),
    ).size,
  );

  function normalize(value: string) {
    return value.trim().toLowerCase();
  }

  // function toggleTheme() {
  //   theme = theme === "dark" ? "light" : "dark";
  // }

  function getPreviewIndex() {
    if (previewTab === "participants") return 1;
    if (previewTab === "allocation") return 2;

    return 0;
  }

  function isExcelFile(file: File) {
    return /\.(xlsx|xls)$/i.test(file.name);
  }

  function formatFileSize(size: number) {
    if (size < 1024) return `${size} B`;

    const sizeInKb = size / 1024;
    if (sizeInKb < 1024) return `${sizeInKb.toFixed(1)} KB`;

    return `${(sizeInKb / 1024).toFixed(1)} MB`;
  }

  function getFileInfo(file: File): UploadedFileInfo {
    return {
      name: file.name,
      size: file.size,
    };
  }

  async function readFirstSheet(file: File) {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer);
    const sheetName = workbook.SheetNames[0];

    if (!sheetName) return null;

    return workbook.Sheets[sheetName] ?? null;
  }

  function getHeaders(sheet: XLSX.WorkSheet) {
    const rows = XLSX.utils.sheet_to_json<string[]>(sheet, {
      header: 1,
      blankrows: false,
    });

    return (rows[0] ?? []).map((header) => String(header).trim());
  }

  function hasExactHeaders(headers: string[], requiredHeaders: string[]) {
    return (
      headers.length === requiredHeaders.length &&
      headers.every((header, index) => header === requiredHeaders[index])
    );
  }

  function resetAllocationState() {
    allocations = [];
    allocated = false;
    allocated_participants = 0;
    unallocated_participants = 0;
    remaining_capacity = total_capacity;
    lastAllocationAt = null;
  }

  function getRoomUsage(room: Room) {
    const occupied = allocations.filter(
      (allocation) =>
        allocation.allottedHostel === room.hostelName &&
        allocation.allottedRoom === room.roomNumber,
    ).length;

    return {
      occupied,
      remaining: room.capacity - occupied,
    };
  }

  function getParticipantStatus(participant: Participant) {
    if (!allocated) return "Pending";

    const allocation = allocations.find(
      (currentAllocation) =>
        currentAllocation.id === participant.id &&
        currentAllocation.gender === participant.gender,
    );

    if (!allocation || allocation.allottedRoom === "-") return "Not Alloted";

    return "Allocated";
  }

  function getRemainingBarColor(remaining: number) {
    if (remaining === 0) return "#EF4444";
    if (remaining === 1) return "#F59E0B";

    return "#22C55E";
  }

  function downloadAllocationExcel() {
    if (!canDownload) return;

    const rows = allocations.map((allocation) => ({
      "Participant Name": allocation.participantName,
      Gender: allocation.gender,
      "Allotted Hostel": allocation.allottedHostel,
      "Allotted Room": allocation.allottedRoom,
    }));
    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, "Allocation Results");
    XLSX.writeFile(workbook, "allocation-results.xlsx");
  }

  async function processHostelFile(file: File) {
    if (!isExcelFile(file)) {
      alert("Please upload a valid Excel file.");
      return;
    }

    try {
      const sheet = await readFirstSheet(file);

      if (!sheet) {
        alert("Hostel file is empty");
        return;
      }

      const rows = XLSX.utils.sheet_to_json<HostelRow>(sheet, {
        defval: "",
      });

      if (rows.length === 0) {
        alert("Hostel file is empty");
        return;
      }

      const requiredHeaders = [
        "Hostel Name",
        "Room Number",
        "Capacity",
        "Hostel Type",
      ];

      if (!hasExactHeaders(getHeaders(sheet), requiredHeaders)) {
        alert("Invalid Hostel Excel format");
        return;
      }

      const parsedRooms = rows.map((row) => ({
        hostelName: String(row["Hostel Name"]).trim(),
        roomNumber: String(row["Room Number"]).trim(),
        capacity: Number(row["Capacity"]),
        hostelType: String(row["Hostel Type"]).trim(),
      }));

      if (
        parsedRooms.length === 0 ||
        parsedRooms.some(
          (room) =>
            !room.hostelName ||
            !room.roomNumber ||
            !room.hostelType ||
            !Number.isFinite(room.capacity) ||
            room.capacity <= 0,
        )
      ) {
        alert("Hostel file contains invalid room details");
        return;
      }

      rooms = parsedRooms;
      total_capacity = rooms.reduce((sum, room) => sum + room.capacity, 0);
      hostelFileInfo = getFileInfo(file);
      resetAllocationState();
      previewTab = "rooms";
    } catch {
      alert("Could not read the hostel Excel file");
    }
  }

  async function processParticipantFile(file: File) {
    if (!isExcelFile(file)) {
      alert("Please upload a valid Excel file.");
      return;
    }

    try {
      const sheet = await readFirstSheet(file);

      if (!sheet) {
        alert("Participant file is empty");
        return;
      }

      const rows = XLSX.utils.sheet_to_json<ParticipantRow>(sheet, {
        defval: "",
      });

      if (rows.length === 0) {
        alert("Participant file is empty");
        return;
      }

      const requiredHeaders = ["Participant Name", "Gender"];

      if (!hasExactHeaders(getHeaders(sheet), requiredHeaders)) {
        alert("Invalid Participant Excel format");
        return;
      }

      const parsedParticipants = rows.map((row, index) => ({
        id: index + 1,
        name: String(row["Participant Name"]).trim(),
        gender: String(row["Gender"]).trim(),
      }));

      if (
        parsedParticipants.length === 0 ||
        parsedParticipants.some(
          (participant) =>
            !participant.name ||
            !["male", "female"].includes(normalize(participant.gender)),
        )
      ) {
        alert("Participant file contains invalid participant details");
        return;
      }

      participants = parsedParticipants;
      participant_no = participants.length;
      participantFileInfo = getFileInfo(file);
      resetAllocationState();
      previewTab = "participants";
    } catch {
      alert("Could not read the participant Excel file");
    }
  }

  async function handleHostelFile(event: Event) {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];

    if (file) {
      await processHostelFile(file);
    }

    input.value = "";
  }

  async function handleParticipantFile(event: Event) {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];

    if (file) {
      await processParticipantFile(file);
    }

    input.value = "";
  }

  function openFilePicker(input: HTMLInputElement | null) {
    if (!input) return;

    input.value = "";
    input.click();
  }

  function handleDragOver(event: DragEvent) {
    event.preventDefault();

    if (event.dataTransfer) {
      event.dataTransfer.dropEffect = "copy";
    }
  }

  async function handleHostelDrop(event: DragEvent) {
    event.preventDefault();

    const file = event.dataTransfer?.files?.[0];

    if (file) {
      await processHostelFile(file);
    }
  }

  async function handleParticipantDrop(event: DragEvent) {
    event.preventDefault();

    const file = event.dataTransfer?.files?.[0];

    if (file) {
      await processParticipantFile(file);
    }
  }

  function formatAllocationTime() {
    return new Intl.DateTimeFormat("en-IN", {
      day: "numeric",
      month: "short",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    }).format(new Date());
  }

  function handleAllocateRooms() {
    if (rooms.length === 0) {
      alert("Upload a valid hostel file.");
      return;
    } else if (participants.length === 0) {
      alert("Upload a valid participant file.");
      return;
    }

    allocations = buildAllocations(participants, rooms);
    allocated_participants = allocations.filter(
      (allocation) => allocation.allottedRoom !== "-",
    ).length;
    unallocated_participants = participants.length - allocated_participants;
    remaining_capacity = total_capacity - allocated_participants;
    allocated = true;
    lastAllocationAt = formatAllocationTime();
    previewTab = "allocation";

    if (unallocated_participants !== 0) {
      alert(
        `${unallocated_participants} participants were not allocated rooms`,
      );
    }
  }
</script>

<!-- Navbar -->
<div
  class:light-theme={theme === "light"}
  class="navbar px-5 py-1 bg-[#0B1220] flex justify-between items-center"
>
  <div class="w-fit flex items-center gap-3">
    <Building2 class="w-7 h-7 text-sky-400" />
    <span class="text-xl text-[#F8FAFC] font-semibold"
      >Hostel Allocation System</span
    >
  </div>

  <div>
    <button
      type="button"
      class="theme-btn border border-[#233554] px-2 py-1 rounded"
      aria-label="Toggle theme"
    >
      {#if theme === "dark"}
        <Moon class="w-4 h-4 text-sky-400" />
      {:else}
        <Sun class="w-4 h-4 text-amber-500" />
      {/if}
    </button>
  </div>
</div>

<div
  class:light-theme={theme === "light"}
  class="main-div pb-10 flex flex-col w-full px-10 bg-[#0A192F]"
>
  <!-- Upper Heading -->
  <div class="py-5 flex flex-col justify-center items-center gap-1">
    <span class="text-3xl h-max text-[#F8FAFC] font-semibold"
      >Hostel Allocation <span class="text-[#38BDF8]">Dashboard</span></span
    >
    <span class="text-[#E2E8F0] text-base"
      >Upload Excel files and allocate rooms in seconds</span
    >
  </div>

  <!-- Stats cards -->
  <div class="stats grid grid-cols-4 gap-3">
    <div class="grid-card">
      <div class="card-logo flex items-center">
        <Bed
          class="w-16 h-16 rounded-full bg-[#1E3A8A33]/80 p-4 text-[#3B82F6] border border-[#1E3A8A33]"
        />
      </div>

      <div class="card-info">
        <span class="text-sm">Total Beds</span>
        <span class="text-2xl text-[#E2E8F0] font-bold"
          >{rooms.length ? total_capacity : "-"}</span
        >
        <span class="text-xs">Across all hostels</span>
      </div>
    </div>

    <div class="grid-card">
      <div class="card-logo">
        <Users
          class="w-16 h-16 rounded-full bg-[#5B21B633]/80 p-4 text-[#8B5CF6] border border-[#5B21B633]"
        />
      </div>

      <div class="card-info">
        <span class="text-sm">Participants</span>
        <span class="text-2xl text-[#E2E8F0] font-bold"
          >{participant_no ? participant_no : "-"}</span
        >
        <span class="text-xs">Total Registered</span>
      </div>
    </div>

    <div class="grid-card">
      <div class="card-logo">
        <Checkcircle
          class="w-16 h-16 rounded-full bg-[#16653433]/80 p-4 text-[#22C55E] border border-[#16653433]"
        />
      </div>

      <div class="card-info">
        <span class="text-sm">Allocated</span>
        <span class="text-2xl text-[#E2E8F0] font-bold"
          >{allocated ? allocated_participants : "-"}</span
        >
        <span class="text-xs">Successfully Alloted</span>
      </div>
    </div>

    <div class="grid-card">
      <div class="card-logo">
        <BedDouble
          class="w-16 h-16 rounded-full bg-[#92400E33]/80 p-4 text-[#F59E0B] border border-[#92400E33]"
        />
      </div>

      <div class="card-info">
        <span class="text-sm">Remaining Beds</span>
        <span class="text-2xl text-[#E2E8F0] font-bold"
          >{rooms.length ? remaining_capacity : "-"}</span
        >
        <span class="text-xs">Vacant Beds</span>
      </div>
    </div>
  </div>

  <div class="upload-section mt-3 grid grid-cols-3 gap-3">
    <div class="upload-grid-card">
      <div class="flex items-center gap-2">
        <FileSpreadsheet
          class="w-10 h-10 rounded-xl p-1 bg-[#3B82F626]/80 border border-[#3B82F626] text-[#3B82F6]"
        />
        <div class="text-[#E2E8F0] flex flex-col align-baseline">
          <span class="text-base font-semibold">Hostel Excel File</span>
          <span class="text-xs"
            >Upload file containing hostel and room details</span
          >
        </div>
      </div>

      <div
        class="upload-input py-3 flex flex-col items-center border border-dashed border-gray-500 rounded-lg transform duration-300 hover:scale-102 hover:bg-[#1b4286]"
        role="group"
        aria-label="Hostel file upload"
        ondragenter={handleDragOver}
        ondragover={handleDragOver}
        ondrop={handleHostelDrop}
      >
        {#if hostelFileInfo}
          <Cloudupload class="w-12 h-12 text-[#38BDF8]" />
          <span
            class="text-[#F8FAFC] text-sm text-center break-all px-2"
            title={hostelFileInfo.name}>{hostelFileInfo.name}</span
          >
          <span class="text-[#64748B] align-bottom text-sm"
            >{formatFileSize(hostelFileInfo.size)}</span
          >
          <button
            type="button"
            onclick={() => openFilePicker(hostelFileInput)}
            class="rounded my-1 px-2 py-1 text-[#38BDF8] border border-[#38BDF8] text-sm cursor-pointer"
            >Change File</button
          >
        {:else}
          <Cloudupload class="w-12 h-12 text-[#38BDF8]" />
          <span class="text-[#F8FAFC] text-sm"
            >Drag & Drop your Excel file here</span
          >
          <span class="text-[#F8FAFC] text-sm">or</span>
          <button
            type="button"
            onclick={() => openFilePicker(hostelFileInput)}
            class="rounded my-1 px-2 py-1 text-[#38BDF8] border border-[#38BDF8] text-sm cursor-pointer"
            >Browse Files</button
          >
          <span class="text-[#64748B] align-bottom text-sm"
            >Supports .xlsx files</span
          >
        {/if}
        <input
          type="file"
          accept=".xlsx,.xls"
          onchange={handleHostelFile}
          bind:this={hostelFileInput}
          class="hidden"
        />
      </div>

      <div class="upload-btn"></div>
    </div>

    <div class="upload-grid-card">
      <div class="flex items-center gap-2">
        <FileUser
          class="w-10 h-10 rounded-xl p-1 bg-[#5B21B633]/80 border border-[#5B21B633] text-[#8B5CF6]"
        />
        <div class="text-[#E2E8F0] flex flex-col align-baseline">
          <span class="text-base font-semibold">Participant Excel File</span>
          <span class="text-xs">Upload file containing participant details</span
          >
        </div>
      </div>
      <div
        class="upload-input py-3 flex flex-col items-center border border-dashed border-gray-500 rounded-lg transform duration-300 hover:scale-102 hover:bg-[#1b4286]"
        role="group"
        aria-label="Participant file upload"
        ondragenter={handleDragOver}
        ondragover={handleDragOver}
        ondrop={handleParticipantDrop}
      >
        {#if participantFileInfo}
          <Cloudupload class="w-12 h-12 text-[#8B5CF6]" />
          <span
            class="text-[#F8FAFC] text-sm text-center break-all px-2"
            title={participantFileInfo.name}>{participantFileInfo.name}</span
          >
          <span class="text-[#64748B] align-bottom text-sm"
            >{formatFileSize(participantFileInfo.size)}</span
          >
          <button
            type="button"
            onclick={() => openFilePicker(participantFileInput)}
            class="rounded my-1 px-2 py-1 text-[#38BDF8] border border-[#38BDF8] text-sm cursor-pointer"
            >Change File</button
          >
        {:else}
          <Cloudupload class="w-12 h-12 text-[#8B5CF6]" />
          <span class="text-[#F8FAFC] text-sm"
            >Drag & Drop your Excel file here</span
          >
          <span class="text-[#F8FAFC] text-sm">or</span>
          <button
            type="button"
            onclick={() => openFilePicker(participantFileInput)}
            class="rounded my-1 px-2 py-1 text-[#38BDF8] border border-[#38BDF8] text-sm cursor-pointer"
            >Browse Files</button
          >
          <span class="text-[#64748B] align-bottom text-sm"
            >Supports .xlsx files</span
          >
        {/if}
        <input
          type="file"
          accept=".xlsx,.xls"
          onchange={handleParticipantFile}
          bind:this={participantFileInput}
          class="hidden"
        />
      </div>
      <div class="upload-btn"></div>
    </div>

    <div
      class="allocation-card flex flex-col gap-10 text-center justify-center items-center rounded-lg bg-[#112240] border border-[#233554] transform duration-300 hover:scale-102 hover:shadow-[0_10px_30px_rgba(0,0,0,0.5)]"
    >
      <div>
        <p class="text-[#F8FAFC] text-base font-semibold">Ready to Allocate?</p>
        <p class="text-[#bfc2c6] text-sm">
          Upload both Excel files to enable room allocation
        </p>
      </div>
      <div class="flex flex-col gap-2">
        <button
          class="grow py-2 flex gap-3 justify-center items-center border font-semibold cursor-pointer text-white border-blue-600 bg-blue-500 hover:bg-blue-600 rounded disabled:bg-gray-400/70 disabled:text-gray-800 disabled:border disabled:border-gray-500 disabled:cursor-not-allowed"
          onclick={handleAllocateRooms}
          disabled={allocationDisabled}
        >
          <Users class="w-5 h-5" />
          Allocate Rooms
        </button>
        <p class="text-sm text-[#64748B]">
          {#if allocated}
            Allocation completed successfully
          {:else if hasBothFiles}
            Click the button to allocate rooms
          {:else}
            Please upload both hostel and participant files first
          {/if}
        </p>
      </div>
    </div>
  </div>

  <div
    class="step-tracker mt-3 rounded-lg bg-[#112240] border border-[#233554] px-6 py-4"
  >
    <div class="step-item">
      <span class:step-complete={hasBothFiles} class="step-number">1</span>
      <span class="text-sm text-[#F8FAFC] font-semibold">Upload Files</span>
      <span class="text-xs text-[#94A3B8]">Upload both Excel files</span>
    </div>

    <div class="step-connector" aria-hidden="true">
      <span class:line-complete={hasBothFiles}></span>
    </div>

    <div class="step-item">
      <span class:step-complete={allocated} class="step-number">2</span>
      <span class="text-sm w-fit text-[#F8FAFC] font-semibold"
        >Allocate Rooms</span
      >
      <span class="text-xs text-[#94A3B8]">System allocates rooms</span>
    </div>

    <div class="step-connector" aria-hidden="true">
      <span class:line-complete={canDownload}></span>
    </div>

    <div class="step-item">
      <span class:step-complete={canDownload} class="step-number">3</span>
      <span class="text-sm text-[#F8FAFC] font-semibold">Download Results</span>
      <span class="text-xs text-[#94A3B8]">Download allocation report</span>
    </div>
  </div>

  <div
    class="data-preview flex gap-4 bg-[#112240] py-3 px-5 mt-3 rounded-lg border border-[#233554]"
  >
    <div class="flex flex-col flex-1">
      <div class="flex gap-2 items-center">
        <List class="w-5 h-5 text-[#3B82F6]" />
        <span class="text-base text-[#E2E8F0] font-semibold">Data Preview</span>
      </div>

      {#if !hasUploadedFiles}
        <div
          class="min-h-32 flex flex-col items-center justify-center gap-2 text-center text-[#94A3B8]"
        >
          <FileSpreadsheet
            class="w-12 h-12 rounded-full p-3 bg-[#1E293B] border border-[#233554] text-[#94A3B8]"
          />
          <span class="text-sm text-[#F8FAFC]">No data to display</span>
          <span class="text-xs">Upload Excel files to preview data here</span>
        </div>
      {:else}
        <div
          class="preview-tabs mt-4 border-b border-[#233554]"
          style={`--tab-index: ${getPreviewIndex()};`}
        >
          <span class="tab-indicator" aria-hidden="true"></span>
          <button
            type="button"
            class:active-tab={previewTab === "rooms"}
            class="preview-tab"
            onclick={() => (previewTab = "rooms")}
          >
            <Building2 class="w-4 h-4" />
            Rooms
          </button>
          <button
            type="button"
            class:active-tab={previewTab === "participants"}
            class="preview-tab"
            onclick={() => (previewTab = "participants")}
          >
            <Users class="w-4 h-4" />
            Participants
          </button>
          <button
            type="button"
            class:active-tab={previewTab === "allocation"}
            class="preview-tab"
            onclick={() => (previewTab = "allocation")}
          >
            <Checkcircle class="w-4 h-4" />
            Allocation
          </button>
        </div>

        <div class="overflow-x-auto mt-3 preview-table-shell">
          {#key previewTab}
            <div class="preview-panel">
              {#if previewTab === "rooms"}
                {#if rooms.length}
                  <table class="preview-table">
                    <thead>
                      <tr>
                        <th>Hostel Name</th>
                        <th>Room Number</th>
                        <th>Capacity</th>
                        <th>Occupied</th>
                        <th>Remaining</th>
                        <th>Hostel Type</th>
                      </tr>
                    </thead>
                    <tbody>
                      {#each rooms as room (room.hostelName + room.roomNumber)}
                        {@const usage = getRoomUsage(room)}
                        <tr>
                          <td>{room.hostelName}</td>
                          <td>{room.roomNumber}</td>
                          <td>{room.capacity}</td>
                          <td>{usage.occupied}</td>
                          <td>
                            <div class="flex items-center gap-3">
                              <span
                                class="remaining-bar"
                                style={`--bar-color: ${getRemainingBarColor(usage.remaining)}; --bar-width: ${
                                  room.capacity
                                    ? (usage.remaining / room.capacity) * 100
                                    : 0
                                }%;`}
                              ></span>
                              <span>{usage.remaining}</span>
                            </div>
                          </td>
                          <td>
                            <span
                              class={normalize(room.hostelType) === "boys"
                                ? "type-badge boys-badge"
                                : "type-badge girls-badge"}
                              >{room.hostelType}</span
                            >
                          </td>
                        </tr>
                      {/each}
                    </tbody>
                  </table>
                {:else}
                  <div class="empty-table-text">
                    Upload hostel file to preview rooms
                  </div>
                {/if}
              {:else if previewTab === "participants"}
                {#if participants.length}
                  <table class="preview-table">
                    <thead>
                      <tr>
                        <th>ID</th>
                        <th>Participant Name</th>
                        <th>Gender</th>
                        <th>Status</th>
                      </tr>
                    </thead>
                    <tbody>
                      {#each participants as participant, index (participant.id ?? index)}
                        {@const status = getParticipantStatus(participant)}
                        <tr>
                          <td>{participant.id ?? index + 1}</td>
                          <td>{participant.name}</td>
                          <td>{participant.gender}</td>
                          <td>
                            <span
                              class={status === "Allocated"
                                ? "status-badge status-done"
                                : status === "Pending"
                                  ? "status-badge status-pending"
                                  : "status-badge status-missing"}
                              >{status}</span
                            >
                          </td>
                        </tr>
                      {/each}
                    </tbody>
                  </table>
                {:else}
                  <div class="empty-table-text">
                    Upload participant file to preview participants
                  </div>
                {/if}
              {:else if allocated && allocations.length}
                <table class="preview-table">
                  <thead>
                    <tr>
                      <th>Participant Name</th>
                      <th>Gender</th>
                      <th>Allotted Hostel</th>
                      <th>Allotted Room</th>
                    </tr>
                  </thead>
                  <tbody>
                    {#each allocations as allocation (allocation.id)}
                      <tr>
                        <td>{allocation.participantName}</td>
                        <td>{allocation.gender}</td>
                        <td>{allocation.allottedHostel}</td>
                        <td>{allocation.allottedRoom}</td>
                      </tr>
                    {/each}
                  </tbody>
                </table>
              {:else}
                <div class="empty-table-text">
                  Allocate rooms to preview allocation results
                </div>
              {/if}
            </div>
          {/key}
        </div>
      {/if}
    </div>

    {#if hasUploadedFiles}
      <div
        class="quick-info flex flex-col gap-3 px-4 py-2 align-baseline border border-[#233554] rounded w-1/4"
      >
        <span class="text-base text-[#E2E8F0] font-semibold">Quick Info</span>

        <div class="flex gap-4 items-center">
          <div class="quick-info-logo">
            <Clock
              class="w-8 h-8 text-[#3B82F6] rounded p-1 bg-[#1D4ED840]/80 border border-[#1D4ED840]"
            />
          </div>
          <div class="flex flex-col">
            <span class="text-sm text-[#E2E8F0] font-semibold"
              >Last Allocation</span
            >
            <span class="text-xs text-[#bfc2c6]">{lastAllocationAt ?? "-"}</span
            >
          </div>
        </div>

        <div class="flex gap-4 items-center">
          <div class="quick-info-logo">
            <Building2
              class="w-8 h-8 text-[#22C55E] rounded p-1 bg-[#16A34A40]/80 border border-[#16A34A40]"
            />
          </div>
          <div class="flex flex-col">
            <span class="text-sm text-[#E2E8F0] font-semibold"
              >Total Hostels</span
            >
            <span class="text-xs text-[#bfc2c6]"
              >{rooms.length ? totalHostels : "-"}</span
            >
          </div>
        </div>

        <div class="flex gap-4 items-center">
          <div class="quick-info-logo">
            <Mars
              class="w-8 h-8 text-[#38BDF8] rounded p-1 bg-[#0284C740]/80 border border-[#0284C740]"
            />
          </div>
          <div class="flex flex-col">
            <span class="text-sm text-[#E2E8F0] font-semibold"
              >Boys Hostels</span
            >
            <span class="text-xs text-[#bfc2c6]"
              >{rooms.length ? boysHostels : "-"}</span
            >
          </div>
        </div>

        <div class="flex gap-4 items-center">
          <div class="quick-info-logo">
            <Venus
              class="w-8 h-8 text-[#EC4899] rounded p-1 bg-[#BE185D40]/80 border border-[#BE185D40]"
            />
          </div>
          <div class="flex flex-col">
            <span class="text-sm text-[#E2E8F0] font-semibold"
              >Girls Hostels</span
            >
            <span class="text-xs text-[#bfc2c6]"
              >{rooms.length ? girlsHostels : "-"}</span
            >
          </div>
        </div>

        <div class="download-section flex flex-col items-center gap-2 my-4">
          <button
            type="button"
            class="download-btn"
            disabled={!canDownload}
            onclick={downloadAllocationExcel}
          >
            <Download class="w-5 h-5" />
            Download Allocation Excel
          </button>
          <span class="text-xs text-[#94A3B8]">
            {canDownload
              ? "Export the allocation results to Excel file"
              : "This will be available after allocation"}
          </span>
        </div>
      </div>
    {/if}
  </div>
</div>

<div
  class="footer bg-[#0B1220]/80 w-full py-3 flex text-[#F8FAFC] flex-col justify-center items-center text-sm"
>
  <span>Built by Divyansh Pandey</span>
  <span class="text-xs">Made with Svelte • Tailwind CSS • SheetJS • TypeScript</span>
  <span class="text-xs">Special thanks to ChatGPT, Codex and Gemini for being my debugging partner throughout this project.</span>
  
  <span class="mt-1"
    ><a href="https://github.com/divyansh-coder-git" target="_blank">GitHub</a>
    •
    <a href="https://www.linkedin.com/in/divyansh-pandey-nits/" target="_blank"
      >LinkedIn</a
    >
    •
    <a href="https://instagram.com/divyansh_coder" target="_blank">Instagram</a
    ></span
  >
</div>

<style>
  :global(body) {
    background-color: #0a192f;
    transition: background-color 0.25s ease;
  }

  /* :global(body:has(.main-div.light-theme)) {
    background-color: #eaf2fb;
  } */

  .navbar,
  .main-div,
  .grid-card,
  .upload-grid-card,
  .data-preview,
  .step-tracker,
  .quick-info,
  .preview-table th,
  .preview-table td,
  .download-btn,
  .theme-btn {
    transition:
      background-color 0.25s ease,
      border-color 0.25s ease,
      color 0.25s ease,
      box-shadow 0.25s ease,
      transform 0.25s ease;
  }

  /* .navbar.light-theme {
    background-color: #ffffff !important;
    border-bottom: 1px solid #cbd5e1;
  }

  .navbar.light-theme span {
    color: #0f172a !important;
  }

  .navbar.light-theme .theme-btn {
    background-color: #f8fafc;
    border-color: #cbd5e1;
  }

  .main-div.light-theme {
    background-color: #eaf2fb !important;
  }

  .main-div.light-theme .grid-card,
  .main-div.light-theme .upload-grid-card,
  .main-div.light-theme .allocation-card,
  .main-div.light-theme .data-preview,
  .main-div.light-theme .step-tracker,
  .main-div.light-theme .quick-info,
  .main-div.light-theme .empty-table-text {
    background-color: #ffffff !important;
    border-color: #cbd5e1 !important;
    box-shadow: 0 10px 25px rgba(15, 23, 42, 0.08);
  }

  .main-div.light-theme :global(.text-\[\#F8FAFC\]),
  .main-div.light-theme :global(.text-\[\#E2E8F0\]),
  .main-div.light-theme :global(.text-white) {
    color: #0f172a !important;
  }

  .main-div.light-theme :global(.text-\[\#bfc2c6\]),
  .main-div.light-theme :global(.text-\[\#64748B\]),
  .main-div.light-theme :global(.text-\[\#94A3B8\]) {
    color: #475569 !important;
  } */
  /* 
  .main-div.light-theme .preview-table {
    color: #0f172a;
    border-color: #cbd5e1;
  }

  .main-div.light-theme .preview-table th {
    background-color: #dbeafe;
    color: #0f172a;
  }

  .main-div.light-theme .preview-table td {
    border-color: #cbd5e1;
  }

  .main-div.light-theme .preview-table tr:nth-child(even) td {
    background-color: #f8fafc;
  }

  .main-div.light-theme .preview-tabs,
  .main-div.light-theme .remaining-bar,
  .main-div.light-theme .step-connector {
    border-color: #cbd5e1 !important;
  }

  .main-div.light-theme .remaining-bar {
    background: linear-gradient(
      to right,
      var(--bar-color) var(--bar-width),
      #dbeafe var(--bar-width)
    );
  }

  .main-div.light-theme .download-btn:disabled {
    background-color: #cbd5e1;
    border-color: #cbd5e1;
    color: #64748b;
  } */

  .grid-card {
    display: flex;
    align-items: center;
    gap: 0.75rem;
    padding: 0.75rem 0.75rem;
    border: 1px solid #233445;
    background-color: #112240;
    border-radius: 0.5rem;
  }

  .card-logo {
    display: flex;
    align-items: center;
  }

  .card-info {
    display: flex;
    flex-direction: column;
    color: #d5dadf;
    align-items: baseline;
    text-align: left;
  }

  .upload-grid-card {
    display: flex;
    flex-direction: column;
    justify-content: space-around;
    gap: 1rem;
    padding: 0.5rem 1rem;
    background-color: #112240;
    border: 1px solid #233554;
    border-radius: 0.5rem;
    transition: 0.3s ease-in-out;
  }

  .upload-grid-card:hover {
    transform: scale(1.02);
    box-shadow: 0 10px 30px rgba(0, 0, 0, 0.5);
  }

  .step-item {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 0.25rem;
    text-align: center;
    position: relative;
  }

  .step-number {
    width: 1.75rem;
    height: 1.75rem;
    display: grid;
    place-items: center;
    border: 1px solid #64748b;
    border-radius: 999px;
    color: #cbd5e1;
    font-size: 0.875rem;
    font-weight: 600;
    background-color: #0f172a;
    transition:
      background-color 0.25s ease,
      border-color 0.25s ease,
      color 0.25s ease,
      box-shadow 0.25s ease,
      transform 0.25s ease;
  }

  .step-number.step-complete {
    border-color: #38bdf8;
    color: #f8fafc;
    background-color: #2563eb;
    transform: scale(1.08);
    box-shadow: 0 0 18px rgba(56, 189, 248, 0.35);
  }

  .step-tracker {
    display: grid;
    grid-template-columns:
      minmax(8rem, 1fr) minmax(4rem, 0.8fr) minmax(8rem, 1fr)
      minmax(4rem, 0.8fr) minmax(8rem, 1fr);
    align-items: start;
    gap: 0.75rem;
  }

  .step-connector {
    position: relative;
    height: 2px;
    margin-top: 0.85rem;
    overflow: hidden;
    border-radius: 999px;
    background-color: #334155;
  }

  .step-connector span {
    position: absolute;
    inset: 0;
    width: 0%;
    border-radius: inherit;
    background: linear-gradient(90deg, #2563eb, #38bdf8);
    transition: width 0.65s ease;
  }

  .step-connector span.line-complete {
    width: 100%;
  }

  .main-div.light-theme .step-number {
    background-color: #ffffff;
    border-color: #94a3b8;
    color: #334155;
  }

  .main-div.light-theme .step-number.step-complete {
    background-color: #2563eb;
    border-color: #2563eb;
    color: #ffffff;
  }

  .preview-tab {
    display: flex;
    align-items: center;
    gap: 0.4rem;
    justify-content: center;
    padding: 0.6rem 0.25rem;
    color: #94a3b8;
    cursor: pointer;
    position: relative;
    z-index: 1;
  }

  .preview-tab:hover,
  .preview-tab.active-tab {
    color: #38bdf8;
  }

  .preview-tabs {
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    position: relative;
    overflow: hidden;
  }

  .tab-indicator {
    position: absolute;
    left: 0;
    bottom: 0;
    width: calc(100% / 3);
    height: 2px;
    border-radius: 999px;
    background-color: #38bdf8;
    transform: translateX(calc(var(--tab-index) * 100%));
    transition: transform 0.28s ease;
  }

  .preview-table-shell {
    overflow: hidden;
  }

  .preview-panel {
    animation: tab-panel-in 0.24s ease;
  }

  @keyframes tab-panel-in {
    from {
      opacity: 0;
      transform: translateX(12px);
    }

    to {
      opacity: 1;
      transform: translateX(0);
    }
  }

  .preview-table {
    width: 100%;
    border-collapse: collapse;
    color: #e2e8f0;
    font-size: 0.875rem;
    overflow: hidden;
    border: 1px solid #233554;
    border-radius: 0.5rem;
  }

  .preview-table th,
  .preview-table td {
    padding: 0.65rem 0.85rem;
    text-align: left;
    border-bottom: 1px solid #233554;
  }

  .preview-table th {
    background-color: #0f2238;
    color: #f8fafc;
    font-weight: 600;
  }

  .preview-table tr:nth-child(even) td {
    background-color: #0d1e33;
  }

  .preview-table tr:last-child td {
    border-bottom: 0;
  }

  .remaining-bar {
    width: 6rem;
    height: 0.45rem;
    border-radius: 999px;
    background: linear-gradient(
      to right,
      var(--bar-color) var(--bar-width),
      #1e293b var(--bar-width)
    );
  }

  .type-badge,
  .status-badge {
    display: inline-flex;
    align-items: center;
    width: fit-content;
    padding: 0.15rem 0.55rem;
    border-radius: 0.35rem;
    font-size: 0.75rem;
    font-weight: 600;
  }

  .boys-badge {
    color: #38bdf8;
    background-color: #1d4ed840;
  }

  .girls-badge {
    color: #f472b6;
    background-color: #be185d40;
  }

  .status-done {
    color: #22c55e;
    background-color: #16653440;
  }

  .status-pending {
    color: #f59e0b;
    background-color: #92400e40;
  }

  .status-missing {
    color: #f87171;
    background-color: #991b1b40;
  }

  .empty-table-text {
    min-height: 6rem;
    display: grid;
    place-items: center;
    color: #94a3b8;
    border: 1px dashed #233554;
    border-radius: 0.5rem;
  }

  .download-btn {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 0.5rem;
    width: fit-content;
    padding: 0.65rem 1rem;
    border: 1px solid #16a34a;
    border-radius: 0.5rem;
    background-color: #16a34a;
    color: #f8fafc;
    font-weight: 700;
    cursor: pointer;
    font-size: 0.9rem;
  }

  .download-btn:hover {
    background-color: #13843c;
  }

  .download-btn:disabled {
    border-color: #475569;
    background-color: #334155;
    color: #94a3b8;
    cursor: not-allowed;
  }
</style>
