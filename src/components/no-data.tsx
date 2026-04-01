import { Session } from "next-auth";
import FrontPage from "./front-page";
import { AlertTriangle } from "lucide-react";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "./ui/card";

const NoData = ({ session }: { session: Session | null }) => {
  return (
    <main>
      {session ? (
        <section className="bg-gradient-to-br from-blue-400 via-teal-300 to-pink-400 min-h-[calc(100dvh-64px)]">
          <div className="container py-12 mx-auto">
            <Card className="mx-auto max-w-xl">
              <CardHeader className="text-center">
                <AlertTriangle className="mx-auto w-16 h-16 text-amber-500" />
                <CardTitle className="text-2xl text-gray-800">
                  No Data Available
                </CardTitle>
                <CardDescription className="text-gray-600">
                  Sorry, it was im<strong>POSSIBLE</strong> to find your data.
                </CardDescription>
              </CardHeader>
              <CardContent className="text-center">
                <p>Please contact Jackie at{" "}
                  <a href="mailto:jackie@heypossible.com" className="text-blue-600 underline">
                    jackie@heypossible.com
                  </a>{" "}
                  so she can investigate.</p>
              </CardContent>
            </Card>
          </div>
        </section>
      ) : (
        <FrontPage />
      )}
    </main>
  );
};

export default NoData;
