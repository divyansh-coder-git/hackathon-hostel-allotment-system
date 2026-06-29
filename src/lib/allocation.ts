
import type {
    Participant,
    Room,
    Allocation
} from "./types";


export function allocateRooms(
    participants: Participant[],
    rooms: Room[],
): Allocation[] {
    console.log("Function executed");

    let i = 0;
    const roomsCopy = $derived(rooms.map((room) => ({ ...room })));

    for (const participant of participants) {
        i++;
        let room;
        let requiredHostel = participant.gender === "Male" ? "Boys" : "Girls";
        room = roomsCopy.find((room) => {
            return room.hostelType == requiredHostel && room.capacity > 0;
        });

        if (!room) {
            if (!room) {
                allocations.push({
                    id: i,
                    participantName: participant.name,
                    gender: participant.gender,
                    allottedHostel: "Not Alloted",
                    allottedRoom: "-",
                });
            }
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
    return allocations;
}