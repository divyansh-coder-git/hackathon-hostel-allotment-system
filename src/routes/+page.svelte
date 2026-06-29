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
  let unallocated_participants = $state(0);

  async function handleHostelFile(event: Event) {
    let input = event.target as HTMLInputElement;
    let file = input.files?.[0];

    if (!file) return;

    let buffer = await file?.arrayBuffer();
    let workbook = XLSX.read(buffer);
    let sheet = workbook.Sheets[workbook.SheetNames[0]];

    const rows = XLSX.utils.sheet_to_json(sheet);

    if (!rows) return;

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

    if (!rows) return;

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

<div>
  <div>
    <input onchange={handleHostelFile} type="file" />
    <input onchange={handleParticipantFile} type="file" />
    <button
      disabled={allocated}
      onclick={() => {
        allocateRooms(participants, roomsCopy);
      }}>Allocate Roms</button
    >
  </div>

  {#each roomsCopy as room (room.roomNumber)}
    <div class="flex gap-3 list-none">
      <li>{room.hostelName}</li>
      <li>{room.roomNumber}</li>
      <li>{room.capacity}</li>
      <li>{room.hostelType}</li>
    </div>
  {/each}
  <span>Total Capacity Left: {total_capacity}</span>
  <br /><br />

  {#each participants as participant (participant.id)}
    <div class="flex gap-3 list-none">
      <li>{participant.name}</li>
      <li>{participant.gender}</li>
    </div>
  {/each}

  <br /><br />

  {#each allocations as allocation (allocation.id)}
    <div class="flex gap-3 list-none">
      <li>{allocation.participantName}</li>
      <li>{allocation.allottedHostel}</li>
      <li>{allocation.allottedRoom}</li>
    </div>
  {/each}
</div>
