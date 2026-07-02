
import type { Allocation, Participant, Room } from "./types";

function normalize(value: string) {
    return value.trim().toLowerCase();
}

export function allocateRooms(
    participants: Participant[],
    rooms: Room[],
): Allocation[] {
    const availableRooms = rooms.map((room) => ({ ...room }));

    return participants.map((participant, index) => {
        const requiredHostel =
            normalize(participant.gender) === "male" ? "boys" : "girls";
        const room = availableRooms.find(
            (currentRoom) =>
                normalize(currentRoom.hostelType) === requiredHostel &&
                currentRoom.capacity > 0,
        );

        if (!room) {
            return {
                id: index + 1,
                participantName: participant.name,
                gender: participant.gender,
                allottedHostel: "Not Alloted",
                allottedRoom: "-",
            };
        }

        room.capacity--;

        return {
            id: index + 1,
            participantName: participant.name,
            gender: participant.gender,
            allottedHostel: room.hostelName,
            allottedRoom: room.roomNumber,
        };
    });
}
