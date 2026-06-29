<script lang="ts">
  import * as XLSX from "xlsx";

  interface Room {
    hostelName: string;
    roomNumber: string;
    capacity: number;
    hostelType: "Boys" | "Girls";
  }

  interface Participant {
    id: number;
    name: string;
    gender: string;
  }

  interface Allocation {
    participantName: string;
    gender: "Male" | "Female";
    allottedHostel: string;
    allottedRoom: string;
}

  let rooms = $state<Room[]>([]);
  let participants = $state<Participant[]>([]);

  async function handleHostelFile(event: Event) {
    let input = event.target as HTMLInputElement;
    let file = input.files?.[0];

    if (!file) return;

    let buffer = await file?.arrayBuffer();
    let workbook = XLSX.read(buffer);
    let sheet = workbook.Sheets[workbook.SheetNames[0]];

    const rows = XLSX.utils.sheet_to_json(sheet);

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    rooms = rows.map((row: any) => {
      return {
        hostelName: row["Hostel Name"],
        roomNumber: row["Room Number"],
        capacity: row["Capacity"],
        hostelType: row["Hostel Type"],
      };
    });

    console.log(rooms);
  }

  async function handleParticipantFile(event: Event) {
    let input = event.target as HTMLInputElement;
    let file = input.files?.[0];

    if (!file) return;

    let buffer = await file?.arrayBuffer();
    let workbook = XLSX.read(buffer);
    let sheet = workbook.Sheets[workbook.SheetNames[0]];

    const rows = XLSX.utils.sheet_to_json(sheet);

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
  }



</script>

<div>
  <div>
    <input onchange={handleHostelFile} type="file" />
    <input onchange={handleParticipantFile} type="file" />
  </div>

  {#each rooms as room (room.roomNumber)}
    <div class="flex gap-3 list-none">
      <li>{room.hostelName}</li>
      <li>{room.roomNumber}</li>
      <li>{room.capacity}</li>
      <li>{room.hostelType}</li>
    </div>
  {/each}
  <br><br>

  {#each participants as participant (participant.id)}
    <div class="flex gap-3 list-none">
      <li>{participant.name}</li>
      <li>{participant.gender}</li>
    </div>
  {/each}
</div>
