import * as React from "react";
import { PlusIcon, ClockIcon, BookOpenIcon, MoreHorizontalIcon, XIcon } from "lucide-react";

export interface HeaderProps {}

const Header: React.FC = () => {
  return (
    <header className="flex justify-between items-center py-1 px-2 bg-transparent z-10">
      <div className="flex items-center text-gray-300 font-medium text-xs">
        Cori
      </div>
      <div className="flex items-center gap-2">
        <PlusIcon className="w-3 h-3 text-gray-400 cursor-pointer hover:text-gray-300 transition-colors" />
        <ClockIcon className="w-3 h-3 text-gray-400 cursor-pointer hover:text-gray-300 transition-colors" />
        <BookOpenIcon className="w-3 h-3 text-gray-400 cursor-pointer hover:text-gray-300 transition-colors" />
        <MoreHorizontalIcon className="w-3 h-3 text-gray-400 cursor-pointer hover:text-gray-300 transition-colors" />
        <XIcon className="w-3 h-3 text-gray-400 cursor-pointer hover:text-gray-300 transition-colors" />
      </div>
    </header>
  );
};

export default Header;
