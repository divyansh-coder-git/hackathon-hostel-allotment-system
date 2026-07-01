<script lang="ts">
  import * as XLSX from "xlsx";
  import Building2 from "@lucide/svelte/icons/building-2";
  import Moon from "@lucide/svelte/icons/moon";
  import Bed from "@lucide/svelte/icons/bed";
  import Users from "@lucide/svelte/icons/users-round";
  import Checkcircle from "@lucide/svelte/icons/circle-check-big";
  import BedDouble from "@lucide/svelte/icons/bed-double";
  import Cloudupload from "@lucide/svelte/icons/cloud-upload";
  import FileSpreadsheet from "@lucide/svelte/icons/file-spreadsheet";
  import FileUser from "@lucide/svelte/icons/file-user";
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

  interface Room {
    hostelName: string;
    roomNumber: string;
    capacity: number;
    hostelType: "Boys" | "Girls" | string;
  }

  interface Participant {
    id: number;
    name: string;
    gender: string;
  }

  interface Allocation {
    id: number;
    participantName: string;
    gender: string;
    allottedHostel: string;
    allottedRoom: string;
  }

  let rooms = $state<Room[]>([]);
  let participants = $state<Participant[]>([]);
  let allocations: Allocation[] = $state([]);

  let total_capacity = $state(0);
  let allocated = $state(true);
  let participant_no = $state(0);
  let unallocated_participants = $state(0);

  async function handleHostelFile(event: Event) {
    let input = event.target as HTMLInputElement;
    let file = input.files?.[0];

    if (!file) return;

    let buffer = await file?.arrayBuffer();
    let workbook = XLSX.read(buffer);
    let sheet = workbook.Sheets[workbook.SheetNames[0]];

    const rows = XLSX.utils.sheet_to_json(sheet);

    if (!rows) {
      return;
    } else if (rows.length === 0) {
      alert("Hostel file is empty");
      return;
    }

    const requiredHeaders = [
      "Hostel Name",
      "Room Number",
      "Capacity",
      "Hostel Type",
    ];

    const headers = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
    })[0] as string[];

    const valid =
      headers.length === requiredHeaders.length &&
      headers.every((header, index) => header === requiredHeaders[index]);

    if (!valid) {
      alert("Invalid Hostel Excel format");
      return;
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    rooms = rows.map((row: any) => {
      total_capacity += row["Capacity"];
      return {
        hostelName: row["Hostel Name"],
        roomNumber: row["Room Number"],
        capacity: row["Capacity"],
        hostelType: row["Hostel Type"],
      };
    });

    console.log(rooms);
    allocated = false;
    allocations = [];
  }

  async function handleParticipantFile(event: Event) {
    let input = event.target as HTMLInputElement;
    let file = input.files?.[0];

    if (!file) return;

    let buffer = await file?.arrayBuffer();
    let workbook = XLSX.read(buffer);
    let sheet = workbook.Sheets[workbook.SheetNames[0]];

    const rows = XLSX.utils.sheet_to_json(sheet);

    if (!rows) {
      return;
    } else if (rows.length === 0) {
      alert("Participant file is empty");
      return;
    }

    const requiredHeaders = ["Participant Name", "Gender"];

    const headers = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
    })[0] as string[];

    const valid =
      headers.length === requiredHeaders.length &&
      headers.every((header, index) => header === requiredHeaders[index]);

    if (!valid) {
      alert("Invalid Participant Excel format");
      return;
    }

    console.log(rows);

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    participants = rows.map((row: any, i) => {
      i++;
      return {
        id: i,
        name: row["Participant Name"],
        gender: row["Gender"],
      };
    });

    participant_no = participants.length;

    allocated = false;
    allocations = [];
  }

  const roomsCopy = $derived(rooms);

  function allocateRooms(participants: Participant[], roomsCopy: Room[]) {
    console.log("Function executed");

    if (roomsCopy.length === 0) {
      alert("Upload a valid hostel file.");
      return;
    } else if (participants.length === 0) {
      alert("Upload a valid participant file.");
      return;
    }

    let i = 0;
    for (const participant of participants) {
      i++;
      let room;
      let requiredHostel = participant.gender === "Male" ? "Boys" : "Girls";
      room = roomsCopy.find((room) => {
        return room.hostelType == requiredHostel && room.capacity > 0;
      });

      if (!room) {
        unallocated_participants++;
        allocations.push({
          id: i,
          participantName: participant.name,
          gender: participant.gender,
          allottedHostel: "Not Alloted",
          allottedRoom: "-",
        });
      } else {
        allocations.push({
          id: i,
          participantName: participant.name,
          gender: participant.gender,
          allottedHostel: room.hostelName,
          allottedRoom: room.roomNumber,
        });
        room.capacity--;
        total_capacity--;
      }
    }
    console.log(allocations);
    if (unallocated_participants != 0) {
      alert(
        `${unallocated_participants} participants were not allocated rooms`,
      );
    }
    allocated = true;
    return allocations;
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
          >{total_capacity ? total_capacity : "-"}</span
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
          >{total_capacity ? total_capacity : "-"}</span
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
          >{total_capacity ? total_capacity : "-"}</span
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
      >
        <Cloudupload class="w-12 h-12 text-[#38BDF8]" />
        <span class="text-[#F8FAFC] text-sm"
          >Drag & Drop your Excel file here</span
        >
        <span class="text-[#F8FAFC] text-sm">or</span>
        <button
          class="rounded my-1 px-2 py-1 text-[#38BDF8] border border-[#38BDF8] text-sm cursor-pointer"
          >Browse Files</button
        >
        <span class="text-[#64748B] align-bottom text-sm"
          >Supports .xlsx files</span
        >
        <input type="file" onchange={handleHostelFile} class="hidden" />
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
      >
        <Cloudupload class="w-12 h-12 text-[#8B5CF6]" />
        <span class="text-[#F8FAFC] text-sm"
          >Drag & Drop your Excel file here</span
        >
        <span class="text-[#F8FAFC] text-sm">or</span>
        <button
          class="rounded my-1 px-2 py-1 text-[#38BDF8] border border-[#38BDF8] text-sm cursor-pointer"
          >Browse Files</button
        >
        <span class="text-[#64748B] align-bottom text-sm"
          >Supports .xlsx files</span
        >
        <input type="file" onchange={handleHostelFile} class="hidden" />
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
          onclick={() => allocateRooms(participants, roomsCopy)}
          disabled={allocated}
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

  <!-- <div class="uploaded-data flex gap-5 px-5 py-5">
    <div class="w-1/3">
      <table>
        <thead>
          <tr>
            <th>Hostel Name</th>
            <th>Room Number</th>
            <th>Capacity</th>
            <th>Hostel Type</th>
          </tr>
        </thead>
        <tbody>
          {#each roomsCopy as room (room.roomNumber)}
            <tr>
              <td>{room.hostelName}</td>
              <td>{room.roomNumber}</td>
              <td>{room.capacity}</td>
              <td>{room.hostelType}</td>
            </tr>
          {/each}
        </tbody>
      </table>
    </div>

    <div class="w-1/3">
      <table>
        <thead>
          <tr>
            <th>Participant Name</th>
            <th>Gender</th>
          </tr>
        </thead>
        <tbody>
          {#each participants as participant (participant.id)}
            <tr>
              <td>{participant.name}</td>
              <td>{participant.gender}</td>
            </tr>
          {/each}
        </tbody>
      </table>
    </div>

    <div class="w-1/3">
      <table>
        <thead>
          <tr>
            <th>Participant Name</th>
            <th>Gender</th>
            <th>Alloted Hostel</th>
            <th>Alloted Room</th>
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
    </div>
  </div> -->

  <div></div>
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
