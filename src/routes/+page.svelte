<script lang="ts">
  import * as XLSX from "xlsx";
  import { allocateRooms as buildAllocations } from "$lib/allocation";
  import type { Allocation, Participant, Room } from "$lib/types";
  import Building2 from "@lucide/svelte/icons/building-2";
  import Moon from "@lucide/svelte/icons/moon";
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
  //   import {

  //Cloudupload
  //     Users,
  //     Bed,
  //     Download,
  //     Upload,
  //     FileSpreadsheet,
  //     CheckCircle,
  //     Moon,
  //     UserCheck,
  //     Calendar,
  //     Mars,
  //     Venus
  // } from "lucide-svelte";

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

  let hostelFileInput = $state<HTMLInputElement | null>(null);
  let participantFileInput = $state<HTMLInputElement | null>(null);
  let hostelFileInfo = $state<UploadedFileInfo | null>(null);
  let participantFileInfo = $state<UploadedFileInfo | null>(null);

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

    if (unallocated_participants !== 0) {
      alert(
        `${unallocated_participants} participants were not allocated rooms`,
      );
    }
  }
</script>

<!-- Navbar -->
<div class="navbar px-5 py-1 bg-[#0B1220] flex justify-between items-center">
  <div class="w-fit flex items-center gap-3">
    <Building2 class="w-7 h-7 text-sky-400" />
    <span class="text-xl text-[#F8FAFC] font-semibold"
      >Hostel Allocation System</span
    >
  </div>

  <div>
    <button class="border border-[#233554] px-2 py-1 rounded">
      <Moon class="w-4 h-4 text-sky-400" />
    </button>
  </div>
</div>

<div class="main-div flex flex-col w-full px-10 bg-[#0A192F]">
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
      class="flex flex-col gap-10 text-center justify-center items-center rounded-lg bg-[#112240] border border-[#233554] transform duration-300 hover:scale-102 hover:shadow-[0_10px_30px_rgba(0,0,0,0.5)]"
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
          Please upload both hostel and participant files first
        </p>
      </div>
    </div>
  </div>
  
  <div
    class="data-preview flex gap-2 bg-[#112240] py-3 px-5 mt-3 rounded-lg border border-[#233554]"
  >
    <div class="flex flex-col flex-1">
      <div class="flex gap-2 items-center">
        <List class="w-5 h-5 text-[#3B82F6]" />
        <span class="text-base text-[#E2E8F0] font-semibold">Data Preview</span>
      </div>
      <div class="data-div"></div>
    </div>

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
          <span class="text-xs text-[#bfc2c6]">{lastAllocationAt ?? "-"}</span>
        </div>
      </div>

      <div class="flex gap-4 items-center">
        <div class="quick-info-logo">
          <Building2
            class="w-8 h-8 text-[#22C55E] rounded p-1 bg-[#16A34A40]/80 border border-[#16A34A40]"
          />
        </div>
        <div class="flex flex-col">
          <span class="text-sm text-[#E2E8F0] font-semibold">Total Hostels</span
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
          <span class="text-sm text-[#E2E8F0] font-semibold">Boys Hostels</span>
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
          <span class="text-sm text-[#E2E8F0] font-semibold">Girls Hostels</span
          >
          <span class="text-xs text-[#bfc2c6]"
            >{rooms.length ? girlsHostels : "-"}</span
          >
        </div>
      </div>
    </div>
  </div>
</div>

<style>
  :global(body) {
    background-color: #0a192f;
  }
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
</style>
