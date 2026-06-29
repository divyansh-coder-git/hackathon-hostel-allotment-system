<script lang="ts">
  import * as XLSX from "xlsx";

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
  let allocated = $state(false);
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

<div class="main-div flex flex-col w-full px-10">
  <div class="navbar h-20"></div>

  <div class="stats grid grid-cols-4 gap-4">
    <div class="grid-card border border-black rounded px-4 py-2 text-center">
      Capacity: {total_capacity}
    </div>
    <div class="grid-card border border-black rounded px-4 py-2 text-center">
      Participants: {participant_no}
    </div>
    <div class="grid-card border border-black rounded px-4 py-2 text-center">
      Alloted:
    </div>
    <div class="grid-card border border-black rounded px-4 py-2 text-center">
      Remaining Capacity:
    </div>
  </div>

  <div class="upload-section flex flex-col gap-5 justify-center w-full">
    <div class="upload-div flex gap-5">
      <div id="hostel_upload"
        class="flex flex-col w-1/2 border border-dashed border-gray-500 text-center justify-center px-10 py-15 mt-5 rounded bg-blue-200/40"
      >
        <label
          for="hostel_input"
          class="text-xl font-semibold text-blue-700 italic"
          >Upload Hostel File</label
        >
        <i class="fa-solid fa-file"></i>
        <label
          for="hostel_input"
          class="text-xl font-semibold text-blue-700 italic"
          >Drag or Click</label
        >
        <input
          class="hidden"
          id="hostel_input"
          onchange={handleHostelFile}
          type="file"
        />
      </div>

      <div id="participant_upload"
        class="flex flex-col w-1/2 border border-dashed border-gray-500 text-center justify-center px-10 py-15 mt-5 rounded bg-blue-200/40"
      >
        <label
          for="participant_input"
          class="text-xl font-semibold text-blue-700 italic"
          >Upload Participant File</label
        >

        <i class="fa-solid fa-user"></i>
        <label
          for="hostel_input"
          class="text-xl font-semibold text-blue-700 italic"
          >Drag or Click</label
        >
        <input
          class="hidden"
          id="participant_input"
          onchange={handleParticipantFile}
          type="file"
        />
      </div>
    </div>

    <button
      disabled={allocated}
      onclick={() => allocateRooms(participants, roomsCopy)}
      >Allocate Rooms</button
    >
  </div>

  <div class="uploaded-data flex gap-5 px-5 py-5">
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
      
  </div>
</div>

<style>
  i{
    margin: auto;
    padding: 1rem 0;
    font-size: 1.5rem;
  }
</style>
